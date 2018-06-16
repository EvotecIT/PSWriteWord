[Flags]
  public enum TableOfContentsSwitches
  {
    None = 0 << 0,

    [Description("\\a")]
    A = 1 << 0,

    [Description("\\b")]
    B = 1 << 1,

    [Description("\\c")]
    C = 1 << 2,

    [Description("\\d")]
    D = 1 << 3,

    [Description("\\f")]
    F = 1 << 4,

    [Description("\\h")]
    H = 1 << 5,

    [Description("\\l")]
    L = 1 << 6,

    [Description("\\n")]
    N = 1 << 7,

    [Description("\\o")]
    O = 1 << 8,

    [Description("\\p")]
    P = 1 << 9,

    [Description("\\s")]
    S = 1 << 10,

    [Description("\\t")]
    T = 1 << 11,

    [Description("\\u")]
    U = 1 << 12,

    [Description("\\w")]
    W = 1 << 13,

    [Description("\\x")]
    X = 1 << 14,

    [Description("\\z")]
    Z = 1 << 15
  }