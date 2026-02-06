async function runStep(scriptPath) {
  const file = document.getElementById("fileInput").files[0];
  if (!file) {
    log("❌ 请选择文件");
    return;
  }

  log("读取文件：" + file.name);

  const data = await file.arrayBuffer();
  const wb = XLSX.read(data);
  const ws = wb.Sheets[wb.SheetNames[0]];
  let aoa = XLSX.utils.sheet_to_json(ws, { header: 1 });

  const vSheet = new VirtualSheet(aoa);
  Application.ActiveSheet = vSheet;

  log("加载脚本：" + scriptPath);

  try {
    const module = await import(`../${scriptPath}`);
    module.run(vSheet);
  } catch (err) {
    log("❌ 脚本加载失败：" + err.message);
    return;
  }

  const outWs = XLSX.utils.aoa_to_sheet(vSheet.data);
  const outWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outWb, outWs, "Sheet1");

  const outName = file.name.replace(/\.(xlsx|xls)$/i, "") + "_step1.xlsx";
  XLSX.writeFile(outWb, outName);

  log("📁 已导出：" + outName);
}
