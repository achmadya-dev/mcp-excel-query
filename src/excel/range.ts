export type A1Bounds = {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

export class Range {
  static letterToCol(letter: string): number {
    let num = 0;
    for (let i = 0; i < letter.length; i++) {
      const charCode = letter.toUpperCase().charCodeAt(i);
      if (charCode < 65 || charCode > 90) {
        throw new Error(`Invalid column character in letter: ${letter}`);
      }
      num = num * 26 + (charCode - 64);
    }
    return num;
  }

  static colToLetter(col: number): string {
    let temp = "";
    let c = col;
    while (c > 0) {
      const rem = (c - 1) % 26;
      temp = String.fromCharCode(65 + rem) + temp;
      c = Math.floor((c - 1) / 26);
    }
    return temp || "A";
  }

  static address(row: number, col: number): string {
    return `${Range.colToLetter(col)}${row}`;
  }

  static toA1(bounds: A1Bounds): string {
    const start = Range.address(bounds.startRow, bounds.startCol);
    if (bounds.startRow === bounds.endRow && bounds.startCol === bounds.endCol) {
      return start;
    }
    return `${start}:${Range.address(bounds.endRow, bounds.endCol)}`;
  }

  static parse(rangeStr: string): A1Bounds {
    const trimmed = rangeStr.trim().replace(/\$/g, "");
    const match = trimmed.match(/^([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$/i);
    if (!match) {
      const singleMatch = trimmed.match(/^([A-Z]+)([0-9]+)$/i);
      if (singleMatch) {
        const col = Range.letterToCol(singleMatch[1]);
        const row = parseInt(singleMatch[2], 10);
        return { startRow: row, startCol: col, endRow: row, endCol: col };
      }
      throw new Error(`Invalid A1 range or cell format: "${rangeStr}"`);
    }
    const startCol = Range.letterToCol(match[1]);
    const startRow = parseInt(match[2], 10);
    const endCol = Range.letterToCol(match[3]);
    const endRow = parseInt(match[4], 10);

    return {
      startRow: Math.min(startRow, endRow),
      startCol: Math.min(startCol, endCol),
      endRow: Math.max(startRow, endRow),
      endCol: Math.max(startCol, endCol),
    };
  }

  static between(startCell: string, endCell?: string): A1Bounds {
    if (endCell) {
      const start = Range.parse(startCell);
      const end = Range.parse(endCell);
      return {
        startRow: Math.min(start.startRow, end.startRow),
        startCol: Math.min(start.startCol, end.startCol),
        endRow: Math.max(start.endRow, end.endRow),
        endCol: Math.max(start.endCol, end.endCol),
      };
    }
    return Range.parse(startCell);
  }
}
