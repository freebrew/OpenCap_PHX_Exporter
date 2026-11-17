export interface GaugeReading {
  upDown: number;
  leftRight: number;
  total: number;
}

export function gaugeLabel(reading: GaugeReading): { vertical: string; side: string; total: string } {
  const vertical = reading.upDown >= 0 ? "UP" : "DN";
  const side = reading.leftRight >= 0 ? "RT" : "LT";
  return {
    vertical: `${Math.abs(reading.upDown).toFixed(2)} ${vertical}`,
    side: `${Math.abs(reading.leftRight).toFixed(2)} ${side}`,
    total: reading.total.toFixed(2),
  };
}

export function gaugeSvg(reading: GaugeReading): string {
  const label = gaugeLabel(reading);
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 160 160" width="160" height="160">
    <circle cx="80" cy="80" r="70" fill="#f7f4ee" stroke="#17324d" stroke-width="2"/>
    <circle cx="80" cy="80" r="36" fill="none" stroke="#2f6f4e" stroke-width="8"/>
    <text x="80" y="74" text-anchor="middle" font-family="Segoe UI, sans-serif" font-size="14">${label.vertical}</text>
    <text x="80" y="92" text-anchor="middle" font-family="Segoe UI, sans-serif" font-size="14">${label.side}</text>
    <text x="80" y="112" text-anchor="middle" font-family="Segoe UI, sans-serif" font-size="12">${label.total}</text>
  </svg>`;
}
