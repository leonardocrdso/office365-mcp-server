const FLAG_ENV_VAR = "OFFICE365_MCP_FLAG_PATH_MARKERS";

const FLAG_LABEL = "[MATERIAL DE OUTRA TURMA/PASTA — confira data e conteúdo antes de citar] ";

function readFlagMarkers(): string[] {
  const raw = process.env[FLAG_ENV_VAR];
  if (!raw) return [];
  return raw
    .split(",")
    .map((marker) => marker.trim())
    .filter((marker) => marker.length > 0);
}

function matchesAnyMarker(haystack: string, markers: string[]): boolean {
  const lowerHaystack = haystack.toLowerCase();
  return markers.some((marker) => lowerHaystack.includes(marker.toLowerCase()));
}

export function isFlagGuardActive(): boolean {
  return readFlagMarkers().length > 0;
}

export function isPathFlagged(...context: Array<string | undefined>): boolean {
  const markers = readFlagMarkers();
  if (markers.length === 0) return false;
  const haystack = context.filter((value): value is string => Boolean(value)).join(" / ");
  return matchesAnyMarker(haystack, markers);
}

export function flagName(name: string, ...pathContext: Array<string | undefined>): string {
  return isPathFlagged(name, ...pathContext) ? `${FLAG_LABEL}${name}` : name;
}
