import type { CellPrimitive, FormulaValue } from "../types.js";

export type FormulaNode =
  | { kind: "call"; name: string; args: FormulaNode[] }
  | { kind: "ref"; address: string }
  | { kind: "range"; address: string }
  | { kind: "number"; value: number }
  | { kind: "string"; value: string }
  | { kind: "boolean"; value: boolean }
  | { kind: "binary"; op: string; left: FormulaNode; right: FormulaNode }
  | { kind: "raw"; formula: string };

function nodeToString(node: FormulaNode): string {
  switch (node.kind) {
    case "call":
      return `${node.name}(${node.args.map(nodeToString).join(",")})`;
    case "ref":
    case "range":
      return node.address;
    case "number":
      return String(node.value);
    case "string":
      return `"${node.value.replace(/"/g, '""')}"`;
    case "boolean":
      return node.value ? "TRUE" : "FALSE";
    case "binary":
      return `${nodeToString(node.left)}${node.op}${nodeToString(node.right)}`;
    case "raw":
      return node.formula;
  }
}

function argToNode(arg: FormulaNode | Expr | string | number | boolean): FormulaNode {
  if (arg instanceof Expr) return arg.node;
  if (typeof arg === "number") return { kind: "number", value: arg };
  if (typeof arg === "boolean") return { kind: "boolean", value: arg };
  if (typeof arg === "string") return { kind: "ref", address: arg };
  return arg;
}

export class Expr {
  constructor(readonly node: FormulaNode) {}

  toString(): string {
    return nodeToString(this.node);
  }

  toFormula(result?: CellPrimitive): FormulaValue {
    return { formula: this.toString(), result };
  }

  get formula(): string {
    return this.toString();
  }

  add(other: Expr | string | number): Expr {
    return new Expr({ kind: "binary", op: "+", left: this.node, right: argToNode(other) });
  }

  sub(other: Expr | string | number): Expr {
    return new Expr({ kind: "binary", op: "-", left: this.node, right: argToNode(other) });
  }

  mul(other: Expr | string | number): Expr {
    return new Expr({ kind: "binary", op: "*", left: this.node, right: argToNode(other) });
  }

  div(other: Expr | string | number): Expr {
    return new Expr({ kind: "binary", op: "/", left: this.node, right: argToNode(other) });
  }
}

function call(name: string, ...args: (FormulaNode | Expr | string | number | boolean)[]): Expr {
  return new Expr({ kind: "call", name, args: args.map(argToNode) });
}

export const Formula = {
  ref(address: string): Expr {
    return new Expr({ kind: "ref", address });
  },

  range(address: string): Expr {
    return new Expr({ kind: "range", address });
  },

  num(value: number): Expr {
    return new Expr({ kind: "number", value });
  },

  str(value: string): Expr {
    return new Expr({ kind: "string", value });
  },

  raw(formula: string): Expr {
    return new Expr({ kind: "raw", formula });
  },

  sum(rangeOrExpr: string | Expr): Expr {
    return call(
      "SUM",
      typeof rangeOrExpr === "string" ? { kind: "range", address: rangeOrExpr } : rangeOrExpr.node,
    );
  },

  average(rangeOrExpr: string | Expr): Expr {
    return call(
      "AVERAGE",
      typeof rangeOrExpr === "string" ? { kind: "range", address: rangeOrExpr } : rangeOrExpr.node,
    );
  },

  count(rangeOrExpr: string | Expr): Expr {
    return call(
      "COUNT",
      typeof rangeOrExpr === "string" ? { kind: "range", address: rangeOrExpr } : rangeOrExpr.node,
    );
  },

  max(rangeOrExpr: string | Expr): Expr {
    return call(
      "MAX",
      typeof rangeOrExpr === "string" ? { kind: "range", address: rangeOrExpr } : rangeOrExpr.node,
    );
  },

  min(rangeOrExpr: string | Expr): Expr {
    return call(
      "MIN",
      typeof rangeOrExpr === "string" ? { kind: "range", address: rangeOrExpr } : rangeOrExpr.node,
    );
  },

  if(
    condition: Expr | string,
    ifTrue: Expr | string | number,
    ifFalse: Expr | string | number,
  ): Expr {
    const condNode =
      typeof condition === "string" ? { kind: "raw" as const, formula: condition } : condition.node;
    return new Expr({
      kind: "call",
      name: "IF",
      args: [condNode, argToNode(ifTrue), argToNode(ifFalse)],
    });
  },

  iferror(value: Expr | string | number, fallback: Expr | string | number): Expr {
    return call("IFERROR", argToNode(value), argToNode(fallback));
  },

  vlookup(lookup: Expr | string | number, table: string, colIndex: number, exact = true): Expr {
    return call("VLOOKUP", argToNode(lookup), { kind: "range", address: table }, colIndex, !exact);
  },

  index(array: string, rowNum: number, colNum?: number): Expr {
    const args: FormulaNode[] = [
      { kind: "range", address: array },
      { kind: "number", value: rowNum },
    ];
    if (colNum !== undefined) args.push({ kind: "number", value: colNum });
    return new Expr({ kind: "call", name: "INDEX", args });
  },

  match(lookup: Expr | string | number, array: string, matchType = 0): Expr {
    return call("MATCH", argToNode(lookup), { kind: "range", address: array }, matchType);
  },

  xlookup(
    lookup: Expr | string | number,
    lookupArray: string,
    returnArray: string,
    ifNotFound?: string,
  ): Expr {
    const args: FormulaNode[] = [
      argToNode(lookup),
      { kind: "range", address: lookupArray },
      { kind: "range", address: returnArray },
    ];
    if (ifNotFound !== undefined) args.push({ kind: "string", value: ifNotFound });
    return new Expr({ kind: "call", name: "XLOOKUP", args });
  },

  concat(...parts: (Expr | string | number)[]): Expr {
    return call(
      "CONCAT",
      ...parts.map((p) =>
        typeof p === "string" ? { kind: "string" as const, value: p } : argToNode(p),
      ),
    );
  },

  round(value: Expr | string | number, digits = 0): Expr {
    return call("ROUND", argToNode(value), digits);
  },

  sumif(range: string, criteria: string | number, sumRange?: string): Expr {
    const args: FormulaNode[] = [
      { kind: "range", address: range },
      typeof criteria === "number"
        ? { kind: "number", value: criteria }
        : { kind: "string", value: criteria },
    ];
    if (sumRange) args.push({ kind: "range", address: sumRange });
    return new Expr({ kind: "call", name: "SUMIF", args });
  },

  countif(range: string, criteria: string | number): Expr {
    return call(
      "COUNTIF",
      { kind: "range", address: range },
      typeof criteria === "number"
        ? { kind: "number", value: criteria }
        : { kind: "string", value: criteria },
    );
  },

  add(...operands: (Expr | string | number)[]): Expr {
    if (operands.length < 2) throw new Error("[Formula] add() requires at least 2 operands");
    let result = argToNode(operands[0]);
    for (let i = 1; i < operands.length; i++) {
      result = { kind: "binary", op: "+", left: result, right: argToNode(operands[i]) };
    }
    return new Expr(result);
  },

  subtract(a: Expr | string | number, b: Expr | string | number): Expr {
    return new Expr({ kind: "binary", op: "-", left: argToNode(a), right: argToNode(b) });
  },

  multiply(a: Expr | string | number, b: Expr | string | number): Expr {
    return new Expr({ kind: "binary", op: "*", left: argToNode(a), right: argToNode(b) });
  },

  divide(a: Expr | string | number, b: Expr | string | number): Expr {
    return new Expr({ kind: "binary", op: "/", left: argToNode(a), right: argToNode(b) });
  },

  pctChange(current: Expr | string | number, previous: Expr | string | number): Expr {
    const c = argToNode(current);
    const p = argToNode(previous);
    return new Expr({
      kind: "binary",
      op: "/",
      left: { kind: "raw", formula: `(${nodeToString(c)}-${nodeToString(p)})` },
      right: p,
    });
  },
} as const;
