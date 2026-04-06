"""
Unified CA-ANN LULC Pipeline (Python only)
===========================================
1) Fetch LULC + predictor rasters from Google Earth Engine (GEE)
2) Train ANN (MLPClassifier) on historical transition (t0 -> t1)
3) Hindcast t2 with Cellular Automata + ANN suitability and validate
4) Simulate future LULC (t2 -> tf)
5) Produce visualisation maps, transition matrices, area stats, exports

Study area: Kharagpur, West Bengal, India

Usage:
  python gee_ca_ann_python_pipeline.py ^
    --min-lon 87.20 --min-lat 22.20 --max-lon 87.45 --max-lat 22.45 ^
    --t0 2017 --t1 2020 --t2 2023 --tf 2026 ^
    --outdir outputs
"""

from __future__ import annotations

import argparse
import io
import json
import os
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

import ee
import matplotlib.colors as mcolors
import matplotlib.pyplot as plt
import numpy as np
import rasterio
import requests
from matplotlib.patches import Patch
from rasterio.transform import Affine
from scipy.ndimage import uniform_filter
from sklearn.inspection import permutation_importance
from sklearn.metrics import (
    accuracy_score,
    classification_report,
    cohen_kappa_score,
    confusion_matrix,
    precision_recall_curve,
    roc_auc_score,
    roc_curve,
)
from sklearn.model_selection import train_test_split
from sklearn.neural_network import MLPClassifier
from sklearn.preprocessing import StandardScaler, label_binarize

# ───────────────────────────────────────────────────────────────
# CONSTANTS
# ───────────────────────────────────────────────────────────────

CLASS_NAMES = [
    "water",
    "trees",
    "grass",
    "flooded_vegetation",
    "crops",
    "shrub_and_scrub",
    "built",
    "bare",
    "snow_and_ice",
]
N_CLASSES = len(CLASS_NAMES)

# Dynamic World colour palette (same order as CLASS_NAMES)
DW_PALETTE = [
    "#419BDF",  # 0 water
    "#397D49",  # 1 trees
    "#88B053",  # 2 grass
    "#7A87C6",  # 3 flooded_vegetation
    "#E49635",  # 4 crops
    "#DFC35A",  # 5 shrub_and_scrub
    "#C4281B",  # 6 built
    "#A59B8F",  # 7 bare
    "#B39FE1",  # 8 snow_and_ice
]


@dataclass
class RasterPack:
    array: np.ndarray
    profile: dict


# ───────────────────────────────────────────────────────────────
# CLI
# ───────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="GEE + Python CA-ANN LULC pipeline")
    p.add_argument("--project", type=str, default=None, help="GEE cloud project id")
    p.add_argument("--min-lon", type=float, required=True)
    p.add_argument("--min-lat", type=float, required=True)
    p.add_argument("--max-lon", type=float, required=True)
    p.add_argument("--max-lat", type=float, required=True)
    p.add_argument("--t0", type=int, default=2017, help="Base year 1")
    p.add_argument("--t1", type=int, default=2020, help="Base year 2 (training target)")
    p.add_argument("--t2", type=int, default=2023, help="Known year for hindcast validation")
    p.add_argument("--tf", type=int, default=2026, help="Future simulation target year")
    p.add_argument("--scale", type=int, default=30, help="Pixel resolution in metres")
    p.add_argument("--max-pixels", type=float, default=1e13)
    p.add_argument("--sample-size", type=int, default=200_000)
    p.add_argument("--hidden-layers", type=str, default="128,128,64")
    p.add_argument("--max-iter", type=int, default=500)
    p.add_argument("--ca-neighborhood", type=int, default=5, help="CA kernel size in pixels")
    p.add_argument("--suitability-weight", type=float, default=0.65)
    p.add_argument("--neighborhood-weight", type=float, default=0.30)
    p.add_argument("--stochastic-weight", type=float, default=0.05, help="CA random perturbation weight")
    p.add_argument("--inertia-threshold", type=float, default=0.45)
    p.add_argument("--skip-download", action="store_true", help="Skip GEE download, use existing data")
    p.add_argument("--seed", type=int, default=42)
    p.add_argument("--outdir", type=str, default="outputs")
    return p.parse_args()


# ───────────────────────────────────────────────────────────────
# GEE INITIALISATION & IMAGE BUILDERS
# ───────────────────────────────────────────────────────────────

def initialize_ee(project: str | None) -> None:
    try:
        ee.Initialize(project=project)
    except Exception:
        ee.Authenticate()
        ee.Initialize(project=project)


def make_aoi(args: argparse.Namespace) -> ee.Geometry:
    return ee.Geometry.Rectangle(
        [args.min_lon, args.min_lat, args.max_lon, args.max_lat],
        proj=None, geodesic=False,
    )


def get_dynamic_world_mode(year: int, aoi: ee.Geometry) -> ee.Image:
    start = ee.Date.fromYMD(year, 1, 1)
    end = start.advance(1, "year")
    return (
        ee.ImageCollection("GOOGLE/DYNAMICWORLD/V1")
        .filterBounds(aoi)
        .filterDate(start, end)
        .select("label")
        .mode()
        .rename("lulc")
        .clip(aoi)
        .toInt8()
    )


def mask_s2_clouds(img: ee.Image) -> ee.Image:
    qa = img.select("QA60")
    cloud_bit = 1 << 10
    cirrus_bit = 1 << 11
    mask = qa.bitwiseAnd(cloud_bit).eq(0).And(qa.bitwiseAnd(cirrus_bit).eq(0))
    return img.updateMask(mask).divide(10000)


def get_s2_predictors(year: int, aoi: ee.Geometry) -> ee.Image:
    start = ee.Date.fromYMD(year, 1, 1)
    end = start.advance(1, "year")
    s2 = (
        ee.ImageCollection("COPERNICUS/S2_SR_HARMONIZED")
        .filterBounds(aoi)
        .filterDate(start, end)
        .filter(ee.Filter.lt("CLOUDY_PIXEL_PERCENTAGE", 35))
        .map(mask_s2_clouds)
    )
    comp = s2.median().clip(aoi)

    spectral = comp.select(["B2", "B3", "B4", "B8", "B11", "B12"]).rename(
        ["blue", "green", "red", "nir", "swir1", "swir2"]
    )
    ndvi = comp.normalizedDifference(["B8", "B4"]).rename("ndvi")
    ndbi = comp.normalizedDifference(["B11", "B8"]).rename("ndbi")
    mndwi = comp.normalizedDifference(["B3", "B11"]).rename("mndwi")
    return spectral.addBands([ndvi, ndbi, mndwi]).toFloat()


def get_static_drivers(aoi: ee.Geometry) -> ee.Image:
    dem = ee.Image("USGS/SRTMGL1_003").rename("elevation").clip(aoi)
    terrain = ee.Terrain.products(dem)
    slope = terrain.select("slope").rename("slope")
    aspect = terrain.select("aspect").rename("aspect")

    viirs = (
        ee.ImageCollection("NOAA/VIIRS/DNB/MONTHLY_V1/VCMSLCFG")
        .filterDate("2023-01-01", "2023-12-31")
        .select("avg_rad")
        .median()
        .rename("nightlights")
        .clip(aoi)
    )
    return ee.Image.cat([dem, slope, aspect, viirs]).toFloat()


def make_distance(image_mask: ee.Image, out_name: str, scale: int) -> ee.Image:
    dist = (
        image_mask.Not()
        .fastDistanceTransform(128, "pixels", "squared_euclidean")
        .sqrt()
        .multiply(scale)
        .rename(out_name)
    )
    return dist.toFloat()


def get_dynamic_drivers(lulc: ee.Image, scale: int) -> ee.Image:
    built = lulc.eq(6)
    forest = lulc.eq(1)
    water = lulc.eq(0)
    d_built = make_distance(built, "dist_built", scale)
    d_forest = make_distance(forest, "dist_forest", scale)
    d_water = make_distance(water, "dist_water", scale)
    return ee.Image.cat([d_built, d_forest, d_water]).toFloat()


def get_predictors(year: int, lulc: ee.Image, aoi: ee.Geometry, scale: int) -> ee.Image:
    static = get_static_drivers(aoi)
    s2 = get_s2_predictors(year, aoi)
    dynamic = get_dynamic_drivers(lulc, scale)
    return ee.Image.cat([s2, static, dynamic]).toFloat()


# ───────────────────────────────────────────────────────────────
# DOWNLOAD & RASTER I/O
# ───────────────────────────────────────────────────────────────

def _download_single_tif(
    image: ee.Image, region_json: str, scale: int,
    out_tif: Path, max_pixels: float,
) -> None:
    """Download one ee.Image to a local GeoTIFF."""
    params = {
        "scale": scale,
        "region": region_json,
        "crs": "EPSG:4326",
        "format": "GEO_TIFF",
        "maxPixels": max_pixels,
    }
    url = image.getDownloadURL(params)
    r = requests.get(url, timeout=300)
    r.raise_for_status()

    content_type = r.headers.get("Content-Type", "").lower()
    if "zip" in content_type or r.content[:2] == b"PK":
        with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
            tif_name = [n for n in zf.namelist() if n.lower().endswith(".tif")][0]
            with zf.open(tif_name) as src, open(out_tif, "wb") as dst:
                dst.write(src.read())
    else:
        with open(out_tif, "wb") as f:
            f.write(r.content)


def download_ee_image_tif(
    image: ee.Image, region: ee.Geometry, scale: int,
    out_tif: Path, max_pixels: float, chunk_size: int = 6,
) -> None:
    """Download an ee.Image, chunking bands if needed to stay under GEE size limit."""
    out_tif.parent.mkdir(parents=True, exist_ok=True)
    region_json = json.dumps(region.getInfo())

    band_names = image.bandNames().getInfo()

    if len(band_names) <= chunk_size:
        _download_single_tif(image, region_json, scale, out_tif, max_pixels)
        return

    # Download in chunks and merge locally
    chunks = [band_names[i:i + chunk_size] for i in range(0, len(band_names), chunk_size)]
    chunk_arrays = []
    ref_profile = None

    for ci, chunk_bands in enumerate(chunks):
        tmp_path = out_tif.parent / f"_chunk_{ci}_{out_tif.name}"
        _download_single_tif(
            image.select(chunk_bands), region_json, scale, tmp_path, max_pixels,
        )
        with rasterio.open(tmp_path) as ds:
            chunk_arrays.append(ds.read())
            if ref_profile is None:
                ref_profile = ds.profile
        tmp_path.unlink()

    merged = np.concatenate(chunk_arrays, axis=0)
    ref_profile.update(count=merged.shape[0], dtype=str(merged.dtype))
    with rasterio.open(out_tif, "w", **ref_profile) as ds:
        ds.write(merged)


def read_raster(path: Path) -> RasterPack:
    with rasterio.open(path) as ds:
        arr = ds.read()
        profile = ds.profile
    return RasterPack(array=arr, profile=profile)


def write_raster(path: Path, arr: np.ndarray, profile: dict) -> None:
    p = dict(profile)
    if arr.ndim == 2:
        arr = arr[np.newaxis, ...]
    p.update(count=arr.shape[0], dtype=str(arr.dtype))
    path.parent.mkdir(parents=True, exist_ok=True)
    with rasterio.open(path, "w", **p) as ds:
        ds.write(arr)


# ───────────────────────────────────────────────────────────────
# ANN TRAINING
# ───────────────────────────────────────────────────────────────

def stack_predictors_and_label(
    predictors: np.ndarray, lulc_from: np.ndarray, lulc_to: np.ndarray,
) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
    b, r, c = predictors.shape
    x = predictors.reshape(b, -1).T
    from_cls = lulc_from.reshape(-1, 1)
    y = lulc_to.reshape(-1)
    valid = np.isfinite(x).all(axis=1) & np.isfinite(from_cls[:, 0]) & np.isfinite(y)
    x_full = np.hstack([x, from_cls])[valid]
    y_full = y[valid].astype(np.int32)
    return x_full, y_full, valid


def train_ann(
    x_train: np.ndarray, y_train: np.ndarray,
    hidden_layers: Tuple[int, ...], max_iter: int, seed: int,
) -> Tuple[MLPClassifier, StandardScaler]:
    scaler = StandardScaler()
    x_train_s = scaler.fit_transform(x_train)

    clf = MLPClassifier(
        hidden_layer_sizes=hidden_layers,
        activation="relu",
        solver="adam",
        max_iter=max_iter,
        random_state=seed,
        learning_rate_init=0.001,
        learning_rate="adaptive",
        early_stopping=True,
        validation_fraction=0.15,
        n_iter_no_change=20,
        batch_size=min(1024, x_train_s.shape[0]),
        verbose=False,
    )
    clf.fit(x_train_s, y_train)

    print(f"  ANN converged at iteration: {clf.n_iter_}")
    if clf.best_loss_ is not None:
        print(f"  Best validation loss: {clf.best_loss_:.6f}")
    else:
        print(f"  Final training loss: {clf.loss_:.6f}")
    return clf, scaler


# ───────────────────────────────────────────────────────────────
# CA SIMULATION
# ───────────────────────────────────────────────────────────────

def neighborhood_fraction_map(lulc: np.ndarray, n_classes: int, kernel_size: int) -> np.ndarray:
    rows, cols = lulc.shape
    out = np.zeros((n_classes, rows, cols), dtype=np.float32)
    for k in range(n_classes):
        mask = (lulc == k).astype(np.float32)
        out[k] = uniform_filter(mask, size=kernel_size, mode="nearest")
    return out


def predict_probabilities(
    clf: MLPClassifier, scaler: StandardScaler,
    predictors: np.ndarray, current_lulc: np.ndarray,
) -> np.ndarray:
    b, r, c = predictors.shape
    x = predictors.reshape(b, -1).T
    x = np.hstack([x, current_lulc.reshape(-1, 1)])
    valid = np.isfinite(x).all(axis=1)

    probs = np.zeros((N_CLASSES, r * c), dtype=np.float32)
    if valid.any():
        x_scaled = scaler.transform(x[valid])
        p = clf.predict_proba(x_scaled)
        clf_classes = clf.classes_.astype(int)
        for i, cls in enumerate(clf_classes):
            probs[cls, valid] = p[:, i]
    return probs.reshape(N_CLASSES, r, c)


def ca_step(
    current_lulc: np.ndarray,
    suitability_probs: np.ndarray,
    suitability_weight: float,
    neighborhood_weight: float,
    stochastic_weight: float,
    inertia_threshold: float,
    kernel_size: int,
) -> np.ndarray:
    rows, cols = current_lulc.shape
    nb = neighborhood_fraction_map(current_lulc, N_CLASSES, kernel_size)

    # Stochastic perturbation for realistic spatial variation
    noise = np.random.uniform(0, 1, size=(N_CLASSES, rows, cols)).astype(np.float32)
    noise = -np.log(-np.log(noise + 1e-10) + 1e-10)  # Gumbel noise
    noise = noise / (noise.max(axis=0, keepdims=True) + 1e-10)

    scores = (
        suitability_weight * suitability_probs
        + neighborhood_weight * nb
        + stochastic_weight * noise
    )
    next_cls = np.argmax(scores, axis=0).astype(np.int16)
    confidence = np.max(scores, axis=0)
    next_cls = np.where(confidence < inertia_threshold, current_lulc, next_cls)
    return next_cls.astype(np.int16)


# ───────────────────────────────────────────────────────────────
# VALIDATION & METRICS
# ───────────────────────────────────────────────────────────────

def evaluate(y_true: np.ndarray, y_pred: np.ndarray, prefix: str) -> Dict[str, object]:
    yt = y_true.reshape(-1)
    yp = y_pred.reshape(-1)
    mask = np.isfinite(yt) & np.isfinite(yp)
    yt = yt[mask].astype(int)
    yp = yp[mask].astype(int)

    acc = float(accuracy_score(yt, yp))
    kappa = float(cohen_kappa_score(yt, yp))
    cm = confusion_matrix(yt, yp, labels=np.arange(N_CLASSES))

    print(f"\n{prefix} Overall Accuracy: {acc:.4f}")
    print(f"{prefix} Kappa: {kappa:.4f}")
    print(f"{prefix} Confusion Matrix (rows=actual, cols=predicted):\n{cm}")

    # Per-class precision / recall / F1
    report = classification_report(
        yt, yp, labels=np.arange(N_CLASSES), target_names=CLASS_NAMES,
        zero_division=0, output_dict=False,
    )
    print(f"\n{prefix} Per-class report:\n{report}")

    report_dict = classification_report(
        yt, yp, labels=np.arange(N_CLASSES), target_names=CLASS_NAMES,
        zero_division=0, output_dict=True,
    )
    return {
        "accuracy": acc, "kappa": kappa,
        "confusion_matrix": cm.tolist(),
        "per_class": report_dict,
    }


def compute_transition_matrix(lulc_from: np.ndarray, lulc_to: np.ndarray) -> np.ndarray:
    f = lulc_from.reshape(-1).astype(int)
    t = lulc_to.reshape(-1).astype(int)
    valid = (f >= 0) & (f < N_CLASSES) & (t >= 0) & (t < N_CLASSES)
    tm = np.zeros((N_CLASSES, N_CLASSES), dtype=np.int64)
    for i in range(N_CLASSES):
        for j in range(N_CLASSES):
            tm[i, j] = int(((f == i) & (t == j) & valid).sum())
    return tm


# ───────────────────────────────────────────────────────────────
# AREA STATISTICS
# ───────────────────────────────────────────────────────────────

def pixel_area_km2(profile: dict) -> float:
    """Compute approximate area of one pixel in km²."""
    transform = profile["transform"]
    crs = profile.get("crs")
    px_w = abs(transform.a)
    px_h = abs(transform.e)

    if crs and hasattr(crs, "is_projected") and crs.is_projected:
        return (px_w * px_h) / 1e6
    else:
        # Geographic CRS (degrees) – approximate using centre latitude
        cy = transform.f + transform.e * profile["height"] / 2
        lat_rad = np.radians(cy)
        m_per_deg_lat = 111_320.0
        m_per_deg_lon = 111_320.0 * np.cos(lat_rad)
        return (px_w * m_per_deg_lon * px_h * m_per_deg_lat) / 1e6


def area_by_class(lulc: np.ndarray, profile: dict) -> Dict[str, float]:
    pix_km2 = pixel_area_km2(profile)
    out = {}
    for k in range(N_CLASSES):
        area = float((lulc == k).sum() * pix_km2)
        out[CLASS_NAMES[k]] = round(area, 4)
    return out


# ───────────────────────────────────────────────────────────────
# VISUALISATION
# ───────────────────────────────────────────────────────────────

def _lulc_cmap():
    """ListedColormap for Dynamic World classes."""
    cmap = mcolors.ListedColormap(DW_PALETTE, N=N_CLASSES)
    norm = mcolors.BoundaryNorm(np.arange(-0.5, N_CLASSES), N_CLASSES)
    return cmap, norm


def plot_lulc_map(lulc: np.ndarray, title: str, out_path: Path) -> None:
    cmap, norm = _lulc_cmap()
    fig, ax = plt.subplots(1, 1, figsize=(10, 8))
    ax.imshow(lulc, cmap=cmap, norm=norm, interpolation="nearest")
    ax.set_title(title, fontsize=14)
    ax.axis("off")
    legend_patches = [
        Patch(facecolor=DW_PALETTE[i], label=CLASS_NAMES[i]) for i in range(N_CLASSES)
    ]
    ax.legend(handles=legend_patches, loc="lower right", fontsize=8, framealpha=0.9)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Map saved: {out_path}")


def plot_change_map(
    lulc_before: np.ndarray, lulc_after: np.ndarray,
    title: str, out_path: Path,
) -> None:
    """Binary change map: red = changed pixel, grey = unchanged."""
    changed = (lulc_before != lulc_after).astype(np.uint8)
    cmap = mcolors.ListedColormap(["#d0d0d0", "#e60000"])
    fig, ax = plt.subplots(1, 1, figsize=(10, 8))
    ax.imshow(changed, cmap=cmap, interpolation="nearest")
    ax.set_title(title, fontsize=14)
    ax.axis("off")
    legend_patches = [
        Patch(facecolor="#d0d0d0", label="No change"),
        Patch(facecolor="#e60000", label="Changed"),
    ]
    ax.legend(handles=legend_patches, loc="lower right", fontsize=9, framealpha=0.9)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Change map saved: {out_path}")


def plot_built_expansion(
    lulc_before: np.ndarray, lulc_after: np.ndarray,
    title: str, out_path: Path,
) -> None:
    """Highlights pixels that became built-up (class 6)."""
    new_built = ((lulc_after == 6) & (lulc_before != 6)).astype(np.uint8)
    kept_built = ((lulc_after == 6) & (lulc_before == 6)).astype(np.uint8)
    rgb = np.full((*lulc_after.shape, 3), 230, dtype=np.uint8)  # light grey bg
    rgb[kept_built == 1] = [200, 60, 30]   # existing built
    rgb[new_built == 1] = [255, 0, 0]      # new built

    fig, ax = plt.subplots(1, 1, figsize=(10, 8))
    ax.imshow(rgb, interpolation="nearest")
    ax.set_title(title, fontsize=14)
    ax.axis("off")
    legend_patches = [
        Patch(facecolor="#c83c1e", label="Existing built-up"),
        Patch(facecolor="#ff0000", label="New built-up"),
        Patch(facecolor="#e6e6e6", label="Other"),
    ]
    ax.legend(handles=legend_patches, loc="lower right", fontsize=9, framealpha=0.9)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Built expansion map saved: {out_path}")


def plot_area_comparison(
    areas: Dict[str, Dict[str, float]], out_path: Path,
) -> None:
    """Grouped bar chart of class areas across time periods."""
    labels = CLASS_NAMES
    periods = list(areas.keys())
    n_periods = len(periods)
    x = np.arange(len(labels))
    width = 0.8 / n_periods

    fig, ax = plt.subplots(figsize=(12, 6))
    for i, period in enumerate(periods):
        vals = [areas[period].get(c, 0) for c in labels]
        bars = ax.bar(x + i * width, vals, width, label=period, color=DW_PALETTE if n_periods == 1 else None)

    ax.set_xlabel("LULC Class")
    ax.set_ylabel("Area (km²)")
    ax.set_title("Class-wise Area Comparison")
    ax.set_xticks(x + width * (n_periods - 1) / 2)
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=8)
    ax.legend()
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Area chart saved: {out_path}")


def plot_ann_loss_curve(clf: MLPClassifier, out_path: Path) -> None:
    """Plot ANN training loss curve over iterations."""
    fig, ax = plt.subplots(figsize=(10, 5))
    epochs = range(1, len(clf.loss_curve_) + 1)
    ax.plot(epochs, clf.loss_curve_, "b-", linewidth=1.5, label="Training loss")

    # Mark early stopping point
    if clf.n_iter_ < len(clf.loss_curve_):
        ax.axvline(x=clf.n_iter_, color="r", linestyle="--", alpha=0.7, label=f"Converged (iter {clf.n_iter_})")

    # Mark best loss
    best_idx = int(np.argmin(clf.loss_curve_))
    best_loss = clf.loss_curve_[best_idx]
    ax.plot(best_idx + 1, best_loss, "r*", markersize=12, label=f"Best loss = {best_loss:.4f}")

    ax.set_xlabel("Iteration", fontsize=12)
    ax.set_ylabel("Loss", fontsize=12)
    ax.set_title("ANN Training Loss Curve", fontsize=14)
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  ANN loss curve saved: {out_path}")


def plot_ann_validation_scores(
    clf: MLPClassifier, out_path: Path,
) -> None:
    """Plot validation score curve if available."""
    if not hasattr(clf, "validation_scores_") or clf.validation_scores_ is None:
        return
    fig, ax = plt.subplots(figsize=(10, 5))
    epochs = range(1, len(clf.validation_scores_) + 1)
    ax.plot(epochs, clf.validation_scores_, "g-", linewidth=1.5, label="Validation accuracy")

    best_idx = int(np.argmax(clf.validation_scores_))
    best_val = clf.validation_scores_[best_idx]
    ax.plot(best_idx + 1, best_val, "r*", markersize=12, label=f"Best = {best_val:.4f}")

    ax.set_xlabel("Iteration", fontsize=12)
    ax.set_ylabel("Validation Accuracy", fontsize=12)
    ax.set_title("ANN Validation Score Curve", fontsize=14)
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  ANN validation curve saved: {out_path}")


def plot_ca_convergence(
    changes_per_step: List[int], total_pixels: int, out_path: Path,
) -> None:
    """Plot CA convergence: pixels changed per step and cumulative change %."""
    steps = range(1, len(changes_per_step) + 1)
    pct = [c / total_pixels * 100 for c in changes_per_step]

    fig, ax1 = plt.subplots(figsize=(10, 5))

    color1 = "#2196F3"
    ax1.bar(steps, changes_per_step, color=color1, alpha=0.7, label="Pixels changed")
    ax1.set_xlabel("CA Step (year)", fontsize=12)
    ax1.set_ylabel("Pixels Changed", fontsize=12, color=color1)
    ax1.tick_params(axis="y", labelcolor=color1)

    ax2 = ax1.twinx()
    color2 = "#FF5722"
    ax2.plot(steps, pct, "o-", color=color2, linewidth=2, markersize=8, label="Change %")
    ax2.set_ylabel("Change (%)", fontsize=12, color=color2)
    ax2.tick_params(axis="y", labelcolor=color2)

    ax1.set_title("CA Simulation Convergence", fontsize=14)
    ax1.set_xticks(list(steps))

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper right", fontsize=10)

    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  CA convergence curve saved: {out_path}")


def plot_roc_curves(
    y_true: np.ndarray, y_proba: np.ndarray, out_path: Path,
) -> None:
    """One-vs-rest ROC curve for each class."""
    y_bin = label_binarize(y_true, classes=np.arange(N_CLASSES))
    fig, ax = plt.subplots(figsize=(10, 8))

    for i in range(N_CLASSES):
        if y_bin[:, i].sum() == 0:
            continue
        fpr, tpr, _ = roc_curve(y_bin[:, i], y_proba[:, i])
        auc = roc_auc_score(y_bin[:, i], y_proba[:, i])
        ax.plot(fpr, tpr, linewidth=1.5, label=f"{CLASS_NAMES[i]} (AUC={auc:.3f})")

    ax.plot([0, 1], [0, 1], "k--", alpha=0.4, linewidth=1)
    ax.set_xlabel("False Positive Rate", fontsize=12)
    ax.set_ylabel("True Positive Rate", fontsize=12)
    ax.set_title("ROC Curves (One-vs-Rest)", fontsize=14)
    ax.legend(fontsize=8, loc="lower right")
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  ROC curves saved: {out_path}")


def plot_precision_recall_curves(
    y_true: np.ndarray, y_proba: np.ndarray, out_path: Path,
) -> None:
    """Precision-Recall curve for each class."""
    y_bin = label_binarize(y_true, classes=np.arange(N_CLASSES))
    fig, ax = plt.subplots(figsize=(10, 8))

    for i in range(N_CLASSES):
        if y_bin[:, i].sum() == 0:
            continue
        prec, rec, _ = precision_recall_curve(y_bin[:, i], y_proba[:, i])
        ax.plot(rec, prec, linewidth=1.5, label=CLASS_NAMES[i])

    ax.set_xlabel("Recall", fontsize=12)
    ax.set_ylabel("Precision", fontsize=12)
    ax.set_title("Precision-Recall Curves (One-vs-Rest)", fontsize=14)
    ax.legend(fontsize=8, loc="lower left")
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  PR curves saved: {out_path}")


def plot_confusion_matrix_heatmap(
    y_true: np.ndarray, y_pred: np.ndarray, title: str, out_path: Path,
) -> None:
    """Normalized confusion matrix as a heatmap."""
    cm = confusion_matrix(y_true, y_pred, labels=np.arange(N_CLASSES))
    # Row-normalize (recall per class)
    row_sums = cm.sum(axis=1, keepdims=True)
    cm_norm = np.divide(cm.astype(float), row_sums, where=row_sums != 0, out=np.zeros_like(cm, dtype=float))

    fig, ax = plt.subplots(figsize=(10, 9))
    im = ax.imshow(cm_norm, cmap="Blues", vmin=0, vmax=1, interpolation="nearest")
    for i in range(N_CLASSES):
        for j in range(N_CLASSES):
            val = cm_norm[i, j]
            color = "white" if val > 0.5 else "black"
            ax.text(j, i, f"{val:.2f}", ha="center", va="center", fontsize=7, color=color)

    ax.set_xticks(range(N_CLASSES))
    ax.set_yticks(range(N_CLASSES))
    ax.set_xticklabels(CLASS_NAMES, rotation=45, ha="right", fontsize=8)
    ax.set_yticklabels(CLASS_NAMES, fontsize=8)
    ax.set_xlabel("Predicted")
    ax.set_ylabel("Actual")
    ax.set_title(title, fontsize=13)
    fig.colorbar(im, ax=ax, label="Recall (row-normalized)")
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Confusion heatmap saved: {out_path}")


def plot_feature_importance(
    clf: MLPClassifier, scaler: StandardScaler,
    x_val: np.ndarray, y_val: np.ndarray,
    feature_names: List[str], out_path: Path,
) -> None:
    """Permutation feature importance on validation set."""
    x_val_s = scaler.transform(x_val)
    result = permutation_importance(
        clf, x_val_s, y_val, n_repeats=5, random_state=42, scoring="accuracy",
    )
    importances = result.importances_mean
    indices = np.argsort(importances)[::-1]

    fig, ax = plt.subplots(figsize=(12, 7))
    n = len(feature_names)
    colors = ["#2196F3" if i < n - 1 else "#FF9800" for i in indices]  # highlight lulc_from
    ax.barh(range(n), importances[indices], color=colors, alpha=0.85)
    ax.set_yticks(range(n))
    ax.set_yticklabels([feature_names[i] for i in indices], fontsize=9)
    ax.invert_yaxis()
    ax.set_xlabel("Mean Accuracy Decrease", fontsize=12)
    ax.set_title("Feature Importance (Permutation)", fontsize=14)
    ax.grid(True, axis="x", alpha=0.3)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Feature importance saved: {out_path}")


def plot_perclass_f1(
    y_true: np.ndarray, y_pred: np.ndarray, title: str, out_path: Path,
) -> None:
    """Bar chart of per-class F1 scores."""
    report = classification_report(
        y_true, y_pred, labels=np.arange(N_CLASSES),
        target_names=CLASS_NAMES, zero_division=0, output_dict=True,
    )
    f1s = [report[c]["f1-score"] for c in CLASS_NAMES]

    fig, ax = plt.subplots(figsize=(11, 6))
    bars = ax.bar(range(N_CLASSES), f1s, color=DW_PALETTE, edgecolor="black", linewidth=0.5)
    for bar, f1 in zip(bars, f1s):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.01,
                f"{f1:.2f}", ha="center", va="bottom", fontsize=9, fontweight="bold")

    ax.set_xticks(range(N_CLASSES))
    ax.set_xticklabels(CLASS_NAMES, rotation=45, ha="right", fontsize=9)
    ax.set_ylabel("F1 Score", fontsize=12)
    ax.set_ylim(0, 1.05)
    ax.set_title(title, fontsize=14)
    ax.axhline(y=np.mean(f1s), color="red", linestyle="--", alpha=0.6, label=f"Mean F1 = {np.mean(f1s):.3f}")
    ax.legend(fontsize=10)
    ax.grid(True, axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Per-class F1 saved: {out_path}")


def plot_spatial_agreement(
    observed: np.ndarray, simulated: np.ndarray,
    title: str, out_path: Path,
) -> None:
    """Map showing where hindcast agrees/disagrees with observed, coloured by error type."""
    correct = (observed == simulated)
    # Commission = simulated says class X but observed is different
    # Omission  = observed is class X but simulated says different
    # Green = correct, Red = wrong
    rgb = np.zeros((*observed.shape, 3), dtype=np.uint8)
    rgb[correct] = [60, 180, 75]      # green = correct
    rgb[~correct] = [230, 25, 25]     # red = wrong

    fig, ax = plt.subplots(figsize=(10, 8))
    ax.imshow(rgb, interpolation="nearest")
    ax.set_title(title, fontsize=14)
    ax.axis("off")
    total = observed.size
    n_correct = int(correct.sum())
    legend_patches = [
        Patch(facecolor="#3cb44b", label=f"Correct ({n_correct/total*100:.1f}%)"),
        Patch(facecolor="#e61919", label=f"Incorrect ({(total-n_correct)/total*100:.1f}%)"),
    ]
    ax.legend(handles=legend_patches, loc="lower right", fontsize=10, framealpha=0.9)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Spatial agreement map saved: {out_path}")


def plot_class_distribution(
    areas: Dict[str, Dict[str, float]], out_path: Path,
) -> None:
    """Stacked area / line chart showing class area evolution over time."""
    periods = list(areas.keys())
    fig, ax = plt.subplots(figsize=(12, 6))

    for i, cls in enumerate(CLASS_NAMES):
        vals = [areas[p].get(cls, 0) for p in periods]
        ax.plot(periods, vals, "o-", color=DW_PALETTE[i], linewidth=2, markersize=8, label=cls)

    ax.set_xlabel("Year", fontsize=12)
    ax.set_ylabel("Area (km²)", fontsize=12)
    ax.set_title("LULC Class Area Trend Over Time", fontsize=14)
    ax.legend(fontsize=8, loc="upper right", ncol=2)
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Class distribution trend saved: {out_path}")


def plot_training_overview(
    clf: MLPClassifier, out_path: Path,
) -> None:
    """Combined loss + validation accuracy on dual axes."""
    fig, ax1 = plt.subplots(figsize=(12, 6))

    # Loss curve
    color1 = "#1565C0"
    epochs = range(1, len(clf.loss_curve_) + 1)
    ax1.plot(epochs, clf.loss_curve_, "-", color=color1, linewidth=1.8, label="Training Loss")
    ax1.set_xlabel("Iteration", fontsize=12)
    ax1.set_ylabel("Loss", fontsize=12, color=color1)
    ax1.tick_params(axis="y", labelcolor=color1)

    # Validation scores on second axis
    if hasattr(clf, "validation_scores_") and clf.validation_scores_ is not None:
        ax2 = ax1.twinx()
        color2 = "#2E7D32"
        ve = range(1, len(clf.validation_scores_) + 1)
        ax2.plot(ve, clf.validation_scores_, "-", color=color2, linewidth=1.8, label="Validation Accuracy")
        ax2.set_ylabel("Validation Accuracy", fontsize=12, color=color2)
        ax2.tick_params(axis="y", labelcolor=color2)

        # Early stopping line
        best_iter = int(np.argmax(clf.validation_scores_)) + 1
        ax1.axvline(x=best_iter, color="red", linestyle="--", alpha=0.6,
                     label=f"Best validation (iter {best_iter})")

        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax1.legend(lines1 + lines2, labels1 + labels2, loc="center right", fontsize=9)
    else:
        ax1.legend(fontsize=10)

    ax1.set_title("ANN Training Overview (Loss + Validation)", fontsize=14)
    ax1.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Training overview saved: {out_path}")


def plot_transition_matrix(tm: np.ndarray, title: str, out_path: Path) -> None:
    """Heatmap of the transition matrix."""
    fig, ax = plt.subplots(figsize=(9, 8))
    im = ax.imshow(tm, cmap="YlOrRd", interpolation="nearest")
    ax.set_xticks(range(N_CLASSES))
    ax.set_yticks(range(N_CLASSES))
    ax.set_xticklabels(CLASS_NAMES, rotation=45, ha="right", fontsize=8)
    ax.set_yticklabels(CLASS_NAMES, fontsize=8)
    ax.set_xlabel("To class")
    ax.set_ylabel("From class")
    ax.set_title(title, fontsize=13)
    fig.colorbar(im, ax=ax, label="Pixel count")
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Transition matrix saved: {out_path}")


def plot_hindcast_comparison(
    observed: np.ndarray, simulated: np.ndarray,
    year: int, out_path: Path,
) -> None:
    """Side-by-side observed vs simulated LULC."""
    cmap, norm = _lulc_cmap()
    fig, axes = plt.subplots(1, 2, figsize=(18, 8))

    axes[0].imshow(observed, cmap=cmap, norm=norm, interpolation="nearest")
    axes[0].set_title(f"Observed {year}", fontsize=13)
    axes[0].axis("off")

    axes[1].imshow(simulated, cmap=cmap, norm=norm, interpolation="nearest")
    axes[1].set_title(f"Simulated {year} (hindcast)", fontsize=13)
    axes[1].axis("off")

    legend_patches = [
        Patch(facecolor=DW_PALETTE[i], label=CLASS_NAMES[i]) for i in range(N_CLASSES)
    ]
    fig.legend(handles=legend_patches, loc="lower center", ncol=N_CLASSES, fontsize=8, framealpha=0.9)
    fig.tight_layout(rect=[0, 0.06, 1, 1])
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"  Hindcast comparison saved: {out_path}")


# ───────────────────────────────────────────────────────────────
# MAIN PIPELINE
# ───────────────────────────────────────────────────────────────

def main() -> None:
    args = parse_args()
    outdir = Path(args.outdir)
    data_dir = outdir / "data"
    maps_dir = outdir / "maps"
    outdir.mkdir(parents=True, exist_ok=True)
    data_dir.mkdir(parents=True, exist_ok=True)
    maps_dir.mkdir(parents=True, exist_ok=True)

    hidden_layers = tuple(int(x.strip()) for x in args.hidden_layers.split(",") if x.strip())
    np.random.seed(args.seed)

    # ── 1–2. GEE: build & download images ───────────────────
    paths = {
        "lulc_t0": data_dir / f"lulc_{args.t0}.tif",
        "lulc_t1": data_dir / f"lulc_{args.t1}.tif",
        "lulc_t2": data_dir / f"lulc_{args.t2}.tif",
        "pred_t0": data_dir / f"predictors_{args.t0}.tif",
        "pred_t1": data_dir / f"predictors_{args.t1}.tif",
        "pred_t2": data_dir / f"predictors_{args.t2}.tif",
    }
    all_exist = all(p.exists() for p in paths.values())

    if args.skip_download and all_exist:
        print("=" * 60)
        print("STEP 1-2: Skipping GEE download (--skip-download, data exists)")
        print("=" * 60)
    else:
        print("=" * 60)
        print("STEP 1: Initialising Google Earth Engine")
        print("=" * 60)
        initialize_ee(args.project)
        aoi = make_aoi(args)

        lulc_t0 = get_dynamic_world_mode(args.t0, aoi)
        lulc_t1 = get_dynamic_world_mode(args.t1, aoi)
        lulc_t2 = get_dynamic_world_mode(args.t2, aoi)

        pred_t0 = get_predictors(args.t0, lulc_t0, aoi, args.scale)
        pred_t1 = get_predictors(args.t1, lulc_t1, aoi, args.scale)
        pred_t2 = get_predictors(args.t2, lulc_t2, aoi, args.scale)

        print("\n" + "=" * 60)
        print("STEP 2: Downloading rasters from GEE")
        print("=" * 60)
        gee_map = {
            "lulc_t0": lulc_t0, "lulc_t1": lulc_t1, "lulc_t2": lulc_t2,
            "pred_t0": pred_t0, "pred_t1": pred_t1, "pred_t2": pred_t2,
        }
        for key, img in gee_map.items():
            print(f"  Downloading {key} ...")
            download_ee_image_tif(img, aoi, args.scale, paths[key], args.max_pixels)
        print("  All downloads complete.")

    # ── 3. Read rasters into numpy ────────────────────────────
    print("\n" + "=" * 60)
    print("STEP 3: Reading rasters")
    print("=" * 60)
    r_lulc_t0 = read_raster(paths["lulc_t0"])
    r_lulc_t1 = read_raster(paths["lulc_t1"])
    r_lulc_t2 = read_raster(paths["lulc_t2"])
    r_pred_t0 = read_raster(paths["pred_t0"])
    r_pred_t1 = read_raster(paths["pred_t1"])
    r_pred_t2 = read_raster(paths["pred_t2"])

    lulc0 = r_lulc_t0.array[0].astype(np.int16)
    lulc1 = r_lulc_t1.array[0].astype(np.int16)
    lulc2 = r_lulc_t2.array[0].astype(np.int16)
    p0 = r_pred_t0.array.astype(np.float32)
    p1 = r_pred_t1.array.astype(np.float32)
    p2 = r_pred_t2.array.astype(np.float32)
    print(f"  LULC shape: {lulc0.shape}  |  Predictor bands: {p0.shape[0]}")

    # ── 4. Train ANN on transition t0 -> t1 ───────────────────
    print("\n" + "=" * 60)
    print("STEP 4: Training ANN (transition t0 -> t1)")
    print("=" * 60)
    x_all, y_all, _ = stack_predictors_and_label(p0, lulc0, lulc1)
    print(f"  Valid training pixels: {x_all.shape[0]}")
    if x_all.shape[0] > args.sample_size:
        idx = np.random.choice(x_all.shape[0], size=args.sample_size, replace=False)
        x_all = x_all[idx]
        y_all = y_all[idx]
        print(f"  Sampled down to: {args.sample_size}")

    x_train, x_val, y_train, y_val = train_test_split(
        x_all, y_all, test_size=0.3, random_state=args.seed, stratify=y_all,
    )
    print(f"  Train: {x_train.shape[0]}  |  Validation: {x_val.shape[0]}")
    print(f"  Hidden layers: {hidden_layers}  |  Max iterations: {args.max_iter}")

    clf, scaler = train_ann(x_train, y_train, hidden_layers, args.max_iter, args.seed)

    x_val_s = scaler.transform(x_val)
    val_pred = clf.predict(x_val_s)
    val_proba = clf.predict_proba(x_val_s)

    # Align proba columns to fixed 0..8 class indices
    val_proba_full = np.zeros((x_val_s.shape[0], N_CLASSES), dtype=np.float32)
    for ci, cls in enumerate(clf.classes_.astype(int)):
        val_proba_full[:, cls] = val_proba[:, ci]

    val_acc = accuracy_score(y_val, val_pred)
    val_kappa = cohen_kappa_score(y_val, val_pred)
    print(f"\n  ANN hold-out validation (t0->t1)")
    print(f"  Accuracy: {val_acc:.4f}")
    print(f"  Kappa:    {val_kappa:.4f}")

    # Per-class report on validation set
    print(classification_report(
        y_val, val_pred, labels=np.arange(N_CLASSES),
        target_names=CLASS_NAMES, zero_division=0,
    ))

    # Feature names for importance plots
    predictor_band_names = [
        "blue", "green", "red", "nir", "swir1", "swir2",
        "ndvi", "ndbi", "mndwi",
        "elevation", "slope", "aspect", "nightlights",
        "dist_built", "dist_forest", "dist_water",
        "lulc_from",
    ]

    # ── 5. Hindcast: simulate t2 from t1 ──────────────────────
    print("\n" + "=" * 60)
    print("STEP 5: Hindcast validation (t1 -> t2)")
    print("=" * 60)
    probs_t1 = predict_probabilities(clf, scaler, p1, lulc1)
    sim_t2 = ca_step(
        current_lulc=lulc1,
        suitability_probs=probs_t1,
        suitability_weight=args.suitability_weight,
        neighborhood_weight=args.neighborhood_weight,
        stochastic_weight=args.stochastic_weight,
        inertia_threshold=args.inertia_threshold,
        kernel_size=args.ca_neighborhood,
    )
    hindcast_metrics = evaluate(lulc2, sim_t2, prefix="Hindcast")

    # ── 6. Future simulation: t2 -> tf ────────────────────────
    print("\n" + "=" * 60)
    print(f"STEP 6: Future simulation (t2 -> tf = {args.tf})")
    print("=" * 60)
    years_ahead = max(1, args.tf - args.t2)
    sim_future = lulc2.copy()
    ca_changes: List[int] = []
    total_pixels = int(lulc2.size)
    for step in range(years_ahead):
        prev = sim_future.copy()
        # Recompute neighbourhood each step for realistic spatial feedback
        nb_probs = predict_probabilities(clf, scaler, p2, sim_future)
        sim_future = ca_step(
            current_lulc=sim_future,
            suitability_probs=nb_probs,
            suitability_weight=args.suitability_weight,
            neighborhood_weight=args.neighborhood_weight,
            stochastic_weight=args.stochastic_weight,
            inertia_threshold=args.inertia_threshold,
            kernel_size=args.ca_neighborhood,
        )
        changed = int((sim_future != prev).sum())
        ca_changes.append(changed)
        print(f"  CA step {step + 1}/{years_ahead} — {changed} pixels changed ({changed/total_pixels*100:.2f}%)")

    # ── 7. Transition matrices ────────────────────────────────
    print("\n" + "=" * 60)
    print("STEP 7: Computing transition matrices")
    print("=" * 60)
    tm_t0_t1 = compute_transition_matrix(lulc0, lulc1)
    tm_t1_t2 = compute_transition_matrix(lulc1, lulc2)
    tm_t2_tf = compute_transition_matrix(lulc2, sim_future)
    print("  Transition matrices computed.")

    # ── 8. Area statistics ────────────────────────────────────
    print("\n" + "=" * 60)
    print("STEP 8: Computing area statistics")
    print("=" * 60)
    ref_profile = r_lulc_t2.profile
    area_t0 = area_by_class(lulc0, r_lulc_t0.profile)
    area_t1 = area_by_class(lulc1, r_lulc_t1.profile)
    area_t2 = area_by_class(lulc2, ref_profile)
    area_future = area_by_class(sim_future, ref_profile)

    print(f"\n  {'Class':<22} {args.t0:>10} {args.t1:>10} {args.t2:>10} {args.tf:>10} (km²)")
    print("  " + "-" * 66)
    for cls in CLASS_NAMES:
        print(
            f"  {cls:<22} "
            f"{area_t0[cls]:>10.2f} {area_t1[cls]:>10.2f} "
            f"{area_t2[cls]:>10.2f} {area_future[cls]:>10.2f}"
        )

    # ── 9. Save rasters ──────────────────────────────────────
    print("\n" + "=" * 60)
    print("STEP 9: Saving output rasters")
    print("=" * 60)
    hindcast_tif = outdir / f"simulated_hindcast_{args.t2}.tif"
    future_tif = outdir / f"simulated_future_{args.tf}.tif"
    write_raster(hindcast_tif, sim_t2.astype(np.int16), ref_profile)
    write_raster(future_tif, sim_future.astype(np.int16), ref_profile)
    print(f"  {hindcast_tif}")
    print(f"  {future_tif}")

    # ── 10. Visualisation ─────────────────────────────────────
    print("\n" + "=" * 60)
    print("STEP 10: Generating visualisation maps")
    print("=" * 60)

    plot_lulc_map(lulc0, f"LULC {args.t0} (observed)", maps_dir / f"lulc_{args.t0}.png")
    plot_lulc_map(lulc1, f"LULC {args.t1} (observed)", maps_dir / f"lulc_{args.t1}.png")
    plot_lulc_map(lulc2, f"LULC {args.t2} (observed)", maps_dir / f"lulc_{args.t2}.png")
    plot_lulc_map(sim_future, f"LULC {args.tf} (simulated)", maps_dir / f"lulc_{args.tf}_simulated.png")

    plot_hindcast_comparison(lulc2, sim_t2, args.t2, maps_dir / f"hindcast_comparison_{args.t2}.png")

    plot_change_map(lulc0, lulc2, f"Change detection {args.t0} → {args.t2}", maps_dir / f"change_{args.t0}_{args.t2}.png")
    plot_change_map(lulc2, sim_future, f"Change detection {args.t2} → {args.tf} (simulated)", maps_dir / f"change_{args.t2}_{args.tf}.png")

    plot_built_expansion(lulc2, sim_future, f"Built-up expansion {args.t2} → {args.tf}", maps_dir / "built_expansion.png")

    plot_transition_matrix(tm_t0_t1, f"Transition matrix {args.t0} → {args.t1}", maps_dir / f"transition_{args.t0}_{args.t1}.png")
    plot_transition_matrix(tm_t1_t2, f"Transition matrix {args.t1} → {args.t2}", maps_dir / f"transition_{args.t1}_{args.t2}.png")
    plot_transition_matrix(tm_t2_tf, f"Transition matrix {args.t2} → {args.tf} (simulated)", maps_dir / f"transition_{args.t2}_{args.tf}.png")

    all_areas = {
        str(args.t0): area_t0,
        str(args.t1): area_t1,
        str(args.t2): area_t2,
        f"{args.tf} (sim)": area_future,
    }
    plot_area_comparison(all_areas, maps_dir / "area_comparison.png")

    # ── ANN training curves ───────────────────────────────────
    plot_ann_loss_curve(clf, maps_dir / "ann_loss_curve.png")
    plot_ann_validation_scores(clf, maps_dir / "ann_validation_curve.png")
    plot_training_overview(clf, maps_dir / "ann_training_overview.png")

    # ── ROC & Precision-Recall curves ─────────────────────────
    plot_roc_curves(y_val, val_proba_full, maps_dir / "roc_curves.png")
    plot_precision_recall_curves(y_val, val_proba_full, maps_dir / "pr_curves.png")

    # ── Confusion matrix heatmaps (normalized) ────────────────
    plot_confusion_matrix_heatmap(
        y_val, val_pred, "ANN Validation Confusion Matrix (normalized)",
        maps_dir / "confusion_ann_validation.png",
    )
    plot_confusion_matrix_heatmap(
        lulc2.reshape(-1).astype(int), sim_t2.reshape(-1).astype(int),
        "Hindcast Confusion Matrix (normalized)",
        maps_dir / "confusion_hindcast.png",
    )

    # ── Per-class F1 bar charts ───────────────────────────────
    plot_perclass_f1(y_val, val_pred, "Per-class F1 — ANN Validation", maps_dir / "f1_ann_validation.png")
    plot_perclass_f1(
        lulc2.reshape(-1).astype(int), sim_t2.reshape(-1).astype(int),
        "Per-class F1 — Hindcast", maps_dir / "f1_hindcast.png",
    )

    # ── Feature importance ────────────────────────────────────
    print("  Computing feature importance (permutation, may take a moment) ...")
    plot_feature_importance(
        clf, scaler, x_val, y_val, predictor_band_names, maps_dir / "feature_importance.png",
    )

    # ── Spatial agreement map ─────────────────────────────────
    plot_spatial_agreement(lulc2, sim_t2, f"Spatial Agreement — Hindcast {args.t2}", maps_dir / "spatial_agreement_hindcast.png")

    # ── Class area trend line ─────────────────────────────────
    plot_class_distribution(all_areas, maps_dir / "class_area_trend.png")

    # ── CA convergence curve ──────────────────────────────────
    if ca_changes:
        plot_ca_convergence(ca_changes, total_pixels, maps_dir / "ca_convergence.png")

    # ── 11. Save JSON summary ─────────────────────────────────
    print("\n" + "=" * 60)
    print("STEP 11: Saving summary JSON")
    print("=" * 60)
    summary = {
        "config": vars(args),
        "classes": {i: n for i, n in enumerate(CLASS_NAMES)},
        "ann_validation": {"accuracy": val_acc, "kappa": val_kappa},
        "hindcast_metrics": hindcast_metrics,
        "transition_matrix_t0_t1": tm_t0_t1.tolist(),
        "transition_matrix_t1_t2": tm_t1_t2.tolist(),
        "transition_matrix_t2_tf": tm_t2_tf.tolist(),
        "area_km2": {
            str(args.t0): area_t0,
            str(args.t1): area_t1,
            str(args.t2): area_t2,
            f"{args.tf}_simulated": area_future,
        },
        "outputs": {
            "hindcast_tif": str(hindcast_tif),
            "future_tif": str(future_tif),
            "maps_directory": str(maps_dir),
        },
    }
    summary_path = outdir / "summary_metrics.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2)
    print(f"  {summary_path}")

    print("\n" + "=" * 60)
    print("PIPELINE COMPLETE")
    print("=" * 60)


if __name__ == "__main__":
    main()
