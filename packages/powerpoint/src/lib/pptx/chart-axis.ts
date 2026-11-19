function extent(values: number[]): [number, number] {
  if (!values.length || values.some((value) => !Number.isFinite(value))) {
    throw new Error("Axis values must be a nonempty array of finite numbers.");
  }
  return [Math.min(...values), Math.max(...values)];
}

export function sharedAxisBounds(values: number[], step = 5, includeZero = true) {
  if (!Number.isFinite(step) || step <= 0) throw new Error("Invalid axis step.");
  const [low, high] = extent(values);
  const min = Math.floor((Math.min(low, includeZero ? 0 : low) - 1) / step) * step;
  const max = Math.ceil((Math.max(high, includeZero ? 0 : high) + 1) / step) * step;
  return { min: min + 0, max: max + 0 };
}

export function centeredAxisBounds(values: number[], minimumSpan = 20) {
  if (!Number.isFinite(minimumSpan) || minimumSpan <= 0) {
    throw new Error("Invalid minimum axis span.");
  }
  const [low, high] = extent(values);
  const midpoint = (low + high) / 2;
  const span = Math.max(minimumSpan, (high - low) * 1.4);
  return { min: midpoint - span / 2, max: midpoint + span / 2 };
}
