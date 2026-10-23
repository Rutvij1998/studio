export type WaveDot = {
  /** 0 to 1 across the slide. */
  x: number;
  /** 0 to 1 down the slide. */
  y: number;
  /** Radius as a fraction of the slide width. The circle stays inside the slide. */
  r: number;
  fill: "FFFFFF" | "FFD6FF" | "FF96FF";
  /** 0 to 1. */
  opacity: number;
};

const SLIDE_W = 1280;
const SLIDE_H = 720;

/**
 * Cover-only particle sheet, matched to the release-showcase still:
 * a perspective mesh, small and faint in the distance, larger and whiter
 * on the crest, with the upper left left clear for the title.
 */
export function coverWave(): WaveDot[] {
  const dots: WaveDot[] = [];
  const cols = 58;
  const rows = 14;
  for (let row = 0; row < rows; row += 1) {
    const z = row / (rows - 1);
    for (let col = 0; col < cols; col += 1) {
      const u = col / (cols - 1);
      const x = (u - 0.4) * 2.2;
      const wave = Math.sin(x * 2.1 + z * 1.9 + 0.2);
      const roll = Math.sin(x * 1.0 - z * 0.45) * 0.42;
      const height = wave * 0.75 + roll;
      const sx = 0.44 + x * (0.32 + z * 0.18);
      const sy = 0.12 + z ** 1.15 * 0.82 + height * (0.05 + z * 0.15) + (1 - u) * 0.04;
      if (sx <= 0 || sx >= 1 || sy <= 0 || sy >= 1) continue;
      // The title lockup owns the top half. One sheet, no stray dots above it.
      if (sy < 0.5) continue;
      const crest = Math.max(0, Math.min(1, (height + 1.15) / 2.3));
      let r = (0.0012 + z ** 1.35 * 0.0068) * (0.62 + crest * 0.5);
      const limit = Math.min(sx, 1 - sx, (sy * SLIDE_H) / SLIDE_W, ((1 - sy) * SLIDE_H) / SLIDE_W);
      if (r > limit) r = limit;
      if (r < 0.00135) continue;
      const opacity = Math.min(0.95, 0.1 + z ** 0.8 * 0.58 + crest * z * 0.28);
      if (opacity < 0.18) continue;
      const fill: WaveDot["fill"] = crest > 0.68 && z > 0.32 ? "FFFFFF" : crest > 0.4 || z > 0.7 ? "FFD6FF" : "FF96FF";
      dots.push({ x: sx, y: sy, r, fill, opacity });
    }
  }
  return dots;
}
