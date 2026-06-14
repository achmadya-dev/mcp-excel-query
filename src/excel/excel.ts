import ExcelJS from "exceljs";
import * as fs from "fs";
import * as path from "path";
import { Range, type A1Bounds } from "./range.js";

export class Sheet {
  constructor(private readonly worksheet: ExcelJS.Worksheet) {}

  get name(): string {
    return this.worksheet.name;
  }

  get raw(): ExcelJS.Worksheet {
    return this.worksheet;
  }

  range(a1: string): A1Bounds;
  range(start: string, end: string): A1Bounds;
  range(startOrA1: string, end?: string): A1Bounds {
    if (end !== undefined) {
      return Range.between(startOrA1, end);
    }
    return Range.parse(startOrA1);
  }
}

export class Excel {
  private constructor(
    private readonly wb: ExcelJS.Workbook,
    private readonly filePath: string
  ) {}

  static pagingLimit(): number {
    const raw = process.env.EXCEL_PAGING_CELLS_LIMIT;
    if (raw === undefined || raw === "") return 4000;
    const n = parseInt(raw, 10);
    return Number.isFinite(n) && n > 0 ? n : 4000;
  }

  static async open(filePath: string): Promise<Excel> {
    if (!fs.existsSync(filePath)) {
      throw new Error(`File not found: ${filePath}`);
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    return new Excel(workbook, filePath);
  }

  static create(filePath: string, opts?: { sheetName?: string; headers?: string[] }): Excel {
    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }

    const workbook = new ExcelJS.Workbook();
    const sheetName = opts?.sheetName ?? "Sheet1";
    const worksheet = workbook.addWorksheet(sheetName);

    if (opts?.headers && opts.headers.length > 0) {
      worksheet.addRow(opts.headers);
    }

    return new Excel(workbook, filePath);
  }

  get path(): string {
    return this.filePath;
  }

  get workbook(): ExcelJS.Workbook {
    return this.wb;
  }

  async save(): Promise<void> {
    await this.wb.xlsx.writeFile(this.filePath);
  }

  sheet(name?: string): Sheet {
    const worksheet = name ? this.wb.getWorksheet(name) : this.wb.worksheets[0];
    if (!worksheet) {
      throw new Error(`Worksheet not found: "${name ?? "Index 0"}"`);
    }
    return new Sheet(worksheet);
  }

  addSheet(name: string): Sheet {
    if (this.wb.getWorksheet(name)) {
      throw new Error(`Worksheet already exists: "${name}"`);
    }
    return new Sheet(this.wb.addWorksheet(name));
  }

  hasSheet(name: string): boolean {
    return this.wb.getWorksheet(name) !== undefined;
  }

  removeSheet(name: string): void {
    if (this.wb.worksheets.length <= 1) {
      throw new Error("Cannot delete the only worksheet in the workbook");
    }
    const worksheet = this.wb.getWorksheet(name);
    if (!worksheet) {
      throw new Error(`Worksheet not found: "${name}"`);
    }
    this.wb.removeWorksheet(worksheet.id);
  }
}
