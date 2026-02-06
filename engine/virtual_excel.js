function log(msg){
  const box = document.getElementById('log');
  box.textContent += msg + "\n";
  box.scrollTop = box.scrollHeight;
}

class VirtualCell {
  constructor(sheet, r, c) {
    this.sheet = sheet;
    this.r = r;
    this.c = c;
  }
  get Value2() {
    return this.sheet.data[this.r-1]?.[this.c-1] ?? null;
  }
  set Value2(v) {
    if (!this.sheet.data[this.r-1]) this.sheet.data[this.r-1] = [];
    this.sheet.data[this.r-1][this.c-1] = v;
  }
  get Text() {
    const v = this.Value2;
    return v == null ? "" : String(v);
  }
  get Interior() {
    const sheet = this.sheet;
    const r = this.r-1, c = this.c-1;
    return {
      set Color(color) {
        if (!sheet.styles[r]) sheet.styles[r] = [];
        if (!sheet.styles[r][c]) sheet.styles[r][c] = {};
        sheet.styles[r][c].bg = color;
      }
    };
  }
  get FormatConditions() {
    return { Delete(){} };
  }
}

class VirtualRow {
  constructor(sheet, r) {
    this.sheet = sheet;
    this.r = r;
  }
  get Interior() {
    const sheet = this.sheet;
    const r = this.r-1;
    return {
      set Color(color) {
        if (!sheet.rowStyles[r]) sheet.rowStyles[r] = {};
        sheet.rowStyles[r].bg = color;
      }
    };
  }
}

class VirtualColumn {
  constructor(sheet, c) {
    this.sheet = sheet;
    this.c = c;
  }
  Insert() {
    const cIndex = this.c-1;
    for (let r = 0; r < this.sheet.data.length; r++) {
      if (!this.sheet.data[r]) this.sheet.data[r] = [];
      this.sheet.data[r].splice(cIndex, 0, null);
    }
  }
}

class VirtualUsedRange {
  constructor(sheet) {
    this.sheet = sheet;
  }
  get Rows() {
    return { Count: this.sheet.data.length };
  }
  get Columns() {
    let max = 0;
    for (let r = 0; r < this.sheet.data.length; r++) {
      max = Math.max(max, this.sheet.data[r]?.length || 0);
    }
    return { Count: max };
  }
}

class VirtualSheet {
  constructor(data) {
    this.data = data;
    this.styles = [];
    this.rowStyles = [];
  }
  Cells(r, c) {
    return new VirtualCell(this, r, c);
  }
  Rows(r) {
    return new VirtualRow(this, r);
  }
  Columns(c) {
    return new VirtualColumn(this, c);
  }
  get UsedRange() {
    return new VirtualUsedRange(this);
  }
  Range(a1) {
    const m = a1.match(/^([A-Z]+)(\d+)$/i);
    if (!m) throw new Error("Range 仅支持如 B2 形式");
    const colLetters = m[1].toUpperCase();
    const row = parseInt(m[2], 10);
    let col = 0;
    for (let i = 0; i < colLetters.length; i++) {
      col = col * 26 + (colLetters.charCodeAt(i) - 64);
    }
    return this.Cells(row, col);
  }
}

const Application = {
  ActiveSheet: null,
  StatusBar: "",
};
