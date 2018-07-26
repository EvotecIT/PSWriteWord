Add-Type -TypeDefinition @"
public enum TableOfContentsSwitches
{
  None = 0 << 0,
  A = 1 << 0,
  B = 1 << 1,
  C = 1 << 2,
  D = 1 << 3,
  F = 1 << 4,
  H = 1 << 5,
  L = 1 << 6,
  N = 1 << 7,
  O = 1 << 8,
  P = 1 << 9,
  S = 1 << 10,
  T = 1 << 11,
  U = 1 << 12,
  W = 1 << 13,
  X = 1 << 14,
  Z = 1 << 15
}
"@