import type {
  ObjectTransform,
  PreviewTransform,
  Rect,
  ScreenPoint,
} from "#/types.ts";

function roundGeometryMetric(value: number): number {
  const rounded = Math.round(value * 1e9) / 1e9;
  return Object.is(rounded, -0) ? 0 : rounded;
}

/** True when a local object or group transform performs no rotation or axis flip. */
export function isIdentityObjectTransform(transform: {
  readonly angle: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}): boolean {
  const normalizedAngle = ((transform.angle % 360) + 360) % 360;
  return (
    (Math.abs(normalizedAngle) < 1e-9 ||
      Math.abs(normalizedAngle - 360) < 1e-9) &&
    !transform.flipH &&
    !transform.flipV
  );
}

/**
 * Computes transformed axis-aligned `bounds`, 4 clockwise `corners`, and own `transform`
 * from an unrotated `box` and an innermost-first transform chain (`[own, ...ancestors]`).
 * Returns `{ bounds: box }` when every step in `transforms` is an identity transform.
 */
export function computeTransformedGeometry(
  box: Rect,
  transforms: readonly PreviewTransform[],
  ownTransform?: PreviewTransform,
): {
  readonly bounds: Rect;
  readonly corners?: readonly [
    ScreenPoint,
    ScreenPoint,
    ScreenPoint,
    ScreenPoint,
  ];
  readonly transform?: ObjectTransform;
} {
  const hasActive = transforms.some((step) => !isIdentityObjectTransform(step));
  if (!hasActive) {
    return { bounds: box };
  }
  const corners: { x: number; y: number }[] = [
    { x: box.x, y: box.y },
    { x: box.x + box.width, y: box.y },
    { x: box.x + box.width, y: box.y + box.height },
    { x: box.x, y: box.y + box.height },
  ];
  for (const step of transforms) {
    const cx = step.box.x + step.box.width / 2;
    const cy = step.box.y + step.box.height / 2;
    const angle = (step.angle * Math.PI) / 180;
    const cos = Math.cos(angle);
    const sin = Math.sin(angle);
    for (const point of corners) {
      const dx = (point.x - cx) * (step.flipH ? -1 : 1);
      const dy = (point.y - cy) * (step.flipV ? -1 : 1);
      point.x = cx + dx * cos - dy * sin;
      point.y = cy + dx * sin + dy * cos;
    }
  }
  const roundedCorners = corners.map((point) => ({
    x: roundGeometryMetric(point.x),
    y: roundGeometryMetric(point.y),
  })) as unknown as readonly [
    ScreenPoint,
    ScreenPoint,
    ScreenPoint,
    ScreenPoint,
  ];
  const left = Math.min(...roundedCorners.map((point) => point.x));
  const top = Math.min(...roundedCorners.map((point) => point.y));
  const right = Math.max(...roundedCorners.map((point) => point.x));
  const bottom = Math.max(...roundedCorners.map((point) => point.y));
  return {
    bounds: {
      x: left,
      y: top,
      width: roundGeometryMetric(right - left),
      height: roundGeometryMetric(bottom - top),
    },
    corners: roundedCorners,
    transform: {
      angle: ownTransform?.angle ?? 0,
      flipH: ownTransform?.flipH ?? false,
      flipV: ownTransform?.flipV ?? false,
    },
  };
}
