import React, { useState, useRef, DragEvent, ChangeEvent, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import localforage from "localforage";
import Tesseract from "tesseract.js";
import {
  Upload,
  Search,
  FileDown,
  Copy,
  Trash2,
  CheckCircle2,
  AlertCircle,
  HelpCircle,
  Database,
  Camera,
  Layers,
  Sparkles,
  CheckSquare,
} from "lucide-react";

interface BatchStockQueryProps {
  systemInventoryMap: Record<string, Record<string, number>>;
  onAutoSwitchToQuery?: (sku: string) => void;
}

export default function BatchStockQuery({
  systemInventoryMap = {},
  onAutoSwitchToQuery,
}: BatchStockQueryProps) {
  // Mode Selection: "system" or "custom"
  const [queryMode, setQueryMode] = useState<"system" | "custom">("system");

  // Shared inputs
  const [skuInput, setSkuInput] = useState<string>("");
  const [ocrLoading, setOcrLoading] = useState<boolean>(false);
  const [matchingMode, setMatchingMode] = useState<"fuzzy" | "exact">("fuzzy");

  // Custom File Mode States
  const [fileName, setFileName] = useState<string>("");
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [currentSheet, setCurrentSheet] = useState<string>("");
  const [customData, setCustomData] = useState<any[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [sumColumn, setSumColumn] = useState<string>("");
  const [skuColumn, setSkuColumn] = useState<string>("");
  const [fileParseStatus, setFileParseStatus] = useState<string>("等待导入文件...");
  const [isSuccessStatus, setIsSuccessStatus] = useState<boolean>(false);
  const [isErrorStatus, setIsErrorStatus] = useState<boolean>(false);

  // Results State
  const [queryResults, setQueryResults] = useState<any[]>([]);
  const [isSearched, setIsSearched] = useState<boolean>(false);
  const [isSearching, setIsSearching] = useState<boolean>(false);
  const [updateNotification, setUpdateNotification] = useState<string>("");

  // Drag Zone State
  const [isDragOver, setIsDragOver] = useState<boolean>(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);

  // Load from localforage on mount
  useEffect(() => {
    const loadStoredData = async () => {
      const storedQueryMode = await localforage.getItem("bsq_queryMode");
      if (storedQueryMode) setQueryMode(storedQueryMode as any);

      const storedSkuInput = await localforage.getItem("bsq_skuInput");
      if (storedSkuInput) setSkuInput(storedSkuInput as any);

      const storedMatchingMode = await localforage.getItem("bsq_matchingMode");
      if (storedMatchingMode) setMatchingMode(storedMatchingMode as any);
      
      const storedFileName = await localforage.getItem("bsq_fileName");
      if (storedFileName) setFileName(storedFileName as any);

      const storedSheetNames = await localforage.getItem("bsq_sheetNames");
      if (storedSheetNames) setSheetNames(storedSheetNames as any);

      const storedCurrentSheet = await localforage.getItem("bsq_currentSheet");
      if (storedCurrentSheet) setCurrentSheet(storedCurrentSheet as any);

      const storedCustomData = await localforage.getItem("bsq_customData");
      if (storedCustomData) {
        setCustomData(storedCustomData as any);
        setIsSuccessStatus(true);
        setFileParseStatus(`成功恢复 ${storedFileName || "文件"} 数据，共 ${(storedCustomData as any).length} 行`);
      }

      const storedHeaders = await localforage.getItem("bsq_headers");
      if (storedHeaders) setHeaders(storedHeaders as any);

      const storedSumColumn = await localforage.getItem("bsq_sumColumn");
      if (storedSumColumn) setSumColumn(storedSumColumn as any);

      const storedSkuColumn = await localforage.getItem("bsq_skuColumn");
      if (storedSkuColumn) setSkuColumn(storedSkuColumn as any);
    };
    loadStoredData();
  }, []);

  useEffect(() => {
    if (queryMode === "custom" && isSearched && skuInput.trim().length > 0 && customData && customData.length > 0) {
      handleQuery();
      setUpdateNotification("上传表格已更新，并已为您自动更新对应结果！");
      setTimeout(() => setUpdateNotification(""), 4000);
    }
  }, [customData]);

  // Save specific states to localforage when they change
  useEffect(() => {
    localforage.setItem("bsq_queryMode", queryMode);
  }, [queryMode]);
  useEffect(() => {
    localforage.setItem("bsq_skuInput", skuInput);
  }, [skuInput]);
  useEffect(() => {
    localforage.setItem("bsq_matchingMode", matchingMode);
  }, [matchingMode]);
  useEffect(() => {
    localforage.setItem("bsq_sumColumn", sumColumn);
  }, [sumColumn]);
  useEffect(() => {
    localforage.setItem("bsq_skuColumn", skuColumn);
  }, [skuColumn]);
  useEffect(() => {
    if (fileName) localforage.setItem("bsq_fileName", fileName);
  }, [fileName]);
  useEffect(() => {
    localforage.setItem("bsq_sheetNames", sheetNames);
  }, [sheetNames]);
  useEffect(() => {
    localforage.setItem("bsq_currentSheet", currentSheet);
  }, [currentSheet]);
  useEffect(() => {
    if (customData.length > 0) localforage.setItem("bsq_customData", customData);
  }, [customData]);
  useEffect(() => {
    if (headers.length > 0) localforage.setItem("bsq_headers", headers);
  }, [headers]);

  // Drag & Drop event handlers
  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = () => {
    setIsDragOver(false);
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragOver(false);
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      processSourceFile(files[0]);
    }
  };

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      processSourceFile(files[0]);
    }
    // Reset file input so the same file can be uploaded again
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  // Helper to load sheets
  const loadExcelSheet = (wb: XLSX.WorkBook, sheetName: string) => {
    try {
      const worksheet = wb.Sheets[sheetName];
      const results: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      if (results.length > 0) {
        setCustomData(results);
        const colHeaders = Object.keys(results[0]);
        setHeaders(colHeaders);

        // Smart Select Sum Column
        const matchSumCol = colHeaders.find(
          (h) =>
            h.includes("可用") ||
            h.includes("库存") ||
            h.includes("当前库存") ||
            h.includes("待发") ||
            h.includes("数") ||
            h.includes("数量")
        ) || colHeaders[0];
        setSumColumn(matchSumCol);

        // Smart Select SKU Column
        const matchSkuCol = colHeaders.find(
          (h) =>
            h.toUpperCase().includes("SKU") ||
            h.includes("编码") ||
            h.includes("商品") ||
            h.includes("品号")
        ) || colHeaders[0];
        setSkuColumn(matchSkuCol);

        setFileParseStatus(`成功加载 [${sheetName}] 表，提供共 ${results.length} 行数据`);
        setIsSuccessStatus(true);
        setIsErrorStatus(false);
        setUpdateNotification("外部本地表格已成功导入并同步！");
        setTimeout(() => setUpdateNotification(""), 4000);
      } else {
        throw new Error("工作表内容为空");
      }
    } catch (err: any) {
      setFileParseStatus(`❌ 表格解析失败: ${err.message}`);
      setIsErrorStatus(true);
      setIsSuccessStatus(false);
      setCustomData([]);
    }
  };

  const handleSheetChange = (e: ChangeEvent<HTMLSelectElement>) => {
    const sheetName = e.target.value;
    setCurrentSheet(sheetName);
    if (workbook) {
      loadExcelSheet(workbook, sheetName);
    }
  };

  // Main file processing logic
  const processSourceFile = (file: File) => {
    const nameLower = file.name.toLowerCase();
    setFileName(file.name);
    setFileParseStatus("正在读取解析文件...");
    setIsSuccessStatus(false);
    setIsErrorStatus(false);

    if (nameLower.endsWith(".csv")) {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          if (results.data && results.data.length > 0) {
            setCustomData(results.data);
            const colHeaders = results.meta.fields || Object.keys(results.data[0]);
            setHeaders(colHeaders);

            // Smart Select Sum Column
            const matchSumCol = colHeaders.find(
              (h) =>
                h.includes("可用") ||
                h.includes("库存") ||
                h.includes("当前库存") ||
                h.includes("待发") ||
                h.includes("数量")
            ) || colHeaders[0];
            setSumColumn(matchSumCol);

            // Smart Select SKU Column
            const matchSkuCol = colHeaders.find(
              (h) =>
                h.toUpperCase().includes("SKU") ||
                h.includes("编码") ||
                h.includes("商品")
            ) || colHeaders[0];
            setSkuColumn(matchSkuCol);

            setFileParseStatus(`成功载入 CSV 文件，共加载了 ${results.data.length} 行数据`);
            setIsSuccessStatus(true);
            setSheetNames([]);
            setCurrentSheet("");
            setWorkbook(null);
            setUpdateNotification("CSV底表已成功导入并同步！");
            setTimeout(() => setUpdateNotification(""), 4000);
          } else {
            setFileParseStatus("❌ CSV 文件的行内容为空");
            setIsErrorStatus(true);
          }
        },
        error: (err) => {
          setFileParseStatus(`❌ CSV 解析失败: ${err.message}`);
          setIsErrorStatus(true);
        },
      });
    } else if (nameLower.endsWith(".xls") || nameLower.endsWith(".xlsx")) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { type: "array" });
          setWorkbook(wb);
          setSheetNames(wb.SheetNames);

          let targetSheet = wb.SheetNames[0];
          // Intelligently find a sheet with "库龄" or "库存" or "Sheet2"
          const preferredSheet = wb.SheetNames.find(
            (name) => name.includes("库龄") || name.includes("库存") || name.includes("Sheet2")
          );
          if (preferredSheet) {
            targetSheet = preferredSheet;
          }

          setCurrentSheet(targetSheet);
          loadExcelSheet(wb, targetSheet);
        } catch (err: any) {
          setFileParseStatus(`❌ Excel 解析失败: ${err.message}`);
          setIsErrorStatus(true);
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setFileParseStatus("❌ 不支持的文件格式，只支持 .csv, .xls, .xlsx 格式文件");
      setIsErrorStatus(true);
    }
  };

  // OCR image text extraction
  const handleImageOcrChange = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setOcrLoading(true);
    try {
      const result = await Tesseract.recognize(file, "eng");
      const text = result.data.text;
      const words = text
        .split(/[\s,，;；\n]+/)
        .map((w) => w.trim())
        .filter((w) => w.length >= 3 && /^[a-zA-Z0-9\-_]+$/.test(wordFormat(w)));

      if (words.length > 0) {
        setSkuInput((prev) => {
          const prevWords = prev
            .split("\n")
            .map((x) => x.trim())
            .filter(Boolean);
          const uniqueNewWords = words.filter((w) => !prevWords.includes(w));
          if (uniqueNewWords.length === 0) return prev;
          const current = prev.trim();
          return current
            ? current + "\n" + uniqueNewWords.join("\n")
            : uniqueNewWords.join("\n");
        });
        alert(`🎉 图片识别成功！发现 ${words.length} 个可能的 SKU，并已自动追加至下方列表中。`);
      } else {
        alert("⚠️ 未从图片中提取到符合标准 SKU 规则 (只含字母/数字/破折号且长度大于2) 的字符串。");
      }
    } catch (err: any) {
      alert("❌ 提取失败: " + err.message);
    } finally {
      setOcrLoading(false);
      if (imageInputRef.current) {
        imageInputRef.current.value = "";
      }
    }
  };

  // Clean-up word help
  const wordFormat = (w: string) => {
    return w.replace(/[^\w-]/g, "");
  };

  // Copy text helper
  const handleCopyResults = () => {
    if (queryResults.length === 0) return;
    const lines = [
      "查询SKU\t汇总数量\t匹配详情/结果"
    ];
    queryResults.forEach(r => {
      lines.push(`${r.querySku}\t${r.stockTotal}\t${r.detailText}`);
    });
    navigator.clipboard.writeText(lines.join("\n"));
    alert("✅ 结果已复制到剪贴板 (可直接粘贴至 Excel 表格中)！");
  };

  // Export results CSV
  const handleExportCSV = () => {
    if (queryResults.length === 0) return;
    let csvContent = "\uFEFF查询SKU,汇总库存,匹配明细/备注\n";
    queryResults.forEach((r) => {
      csvContent += `"${r.querySku}","${r.stockTotal}","${r.detailText}"\n`;
    });

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "SKU批量查询结果汇总.csv");
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Execute querying
  const handleQuery = () => {
    const listSkus = skuInput
      .split("\n")
      .map((s) => s.trim())
      .filter((s) => s.length > 0);

    if (listSkus.length === 0) {
      alert("⚠️ 请先在输入框中输入需查询的 SKU (每行一个)！");
      return;
    }

    setIsSearching(true);

    setTimeout(() => {
      const results: any[] = [];

      if (queryMode === "system") {
        // Direct Query from loaded system inventory Map
        listSkus.forEach((qSku) => {
          let stockTotal = 0;
          const whBreakdowns: string[] = [];
          const qUpper = qSku.toUpperCase();

          // Search inside systemInventoryMap
          // It's structured as Record<sku, Record<whName, stockQty>>
          let foundMatches = false;

          for (const sysSkuKey in systemInventoryMap) {
            const matchesSku = matchingMode === "fuzzy"
              ? sysSkuKey.toUpperCase().includes(qUpper) || qUpper.includes(sysSkuKey.toUpperCase())
              : sysSkuKey.toUpperCase() === qUpper;

            if (matchesSku) {
              foundMatches = true;
              const whs = systemInventoryMap[sysSkuKey];
              for (const whName in whs) {
                const qty = whs[whName] || 0;
                if (qty > 0) {
                  stockTotal += qty;
                  whBreakdowns.push(`${whName}(${qty})`);
                }
              }
            }
          }

          results.push({
            querySku: qSku,
            stockTotal: foundMatches ? stockTotal : 0,
            detailText: foundMatches
              ? whBreakdowns.length > 0
                ? whBreakdowns.join(" | ")
                : "全仓无库存"
              : "库存系统中未匹配到该 SKU",
            matchStatus: foundMatches ? "matched" : "not-found",
          });
        });
      } else {
        // Query from custom uploaded spreadsheet columns
        if (customData.length === 0) {
          alert("⚠️ 请先导入表格数据源！");
          setIsSearching(false);
          return;
        }

        const matchCol = skuColumn || headers[0];
        const valCol = sumColumn;

        listSkus.forEach((qSku) => {
          let stockTotal = 0;
          let matchRowsCount = 0;
          const qUpper = qSku.toUpperCase();

          customData.forEach((row) => {
            const rowSkuStr = String(row[matchCol] || "").trim();
            const rowSkuUpper = rowSkuStr.toUpperCase();

            const isMatch = matchingMode === "fuzzy"
              ? rowSkuUpper.includes(qUpper) || qUpper.includes(rowSkuUpper)
              : rowSkuUpper === qUpper;

            if (rowSkuStr && isMatch) {
              const parseVal = parseFloat(String(row[valCol] || "0").replace(/,/g, ""));
              if (!isNaN(parseVal)) {
                stockTotal += parseVal;
              }
              matchRowsCount++;
            }
          });

          results.push({
            querySku: qSku,
            stockTotal: matchRowsCount > 0 ? stockTotal : 0,
            detailText: matchRowsCount > 0 ? `匹配到 ${matchRowsCount} 行记录` : "无匹配记录",
            matchStatus: matchRowsCount > 0 ? "matched" : "not-found",
          });
        });
      }

      setQueryResults(results);
      setIsSearched(true);
      setIsSearching(false);
      setUpdateNotification("查询处理完成，数据展示已更新！");
      setTimeout(() => setUpdateNotification(""), 3000);
    }, 150);
  };

  const hasSystemData = Object.keys(systemInventoryMap).length > 0;

  // Automatically re-query if system inventory updates and we are in system mode
  React.useEffect(() => {
    if (queryMode === "system" && isSearched && skuInput.trim().length > 0) {
      handleQuery();
      setUpdateNotification("系统底表库存已更新，并已自动重新计算查询结果！");
      setTimeout(() => setUpdateNotification(""), 4000);
    }
  }, [systemInventoryMap]);

  return (
    <div className="space-y-6 relative">
      {/* Toast Notification */}
      {updateNotification && (
        <div className="fixed top-6 left-1/2 -translate-x-1/2 z-50 animate-in fade-in slide-in-from-top-4 duration-300">
          <div className="bg-emerald-600 text-white px-5 py-3 rounded-2xl shadow-xl flex items-center gap-3 border-2 border-emerald-500/30">
            <CheckCircle2 className="w-5 h-5 text-emerald-100" />
            <span className="text-sm font-bold tracking-wide">{updateNotification}</span>
          </div>
        </div>
      )}

      {/* Tab controls to choose system mapping vs custom sheet mapping */}
      <div className="bg-white rounded-xl shadow-sm border border-zinc-200 overflow-hidden">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between p-6 bg-zinc-50/50 border-b border-zinc-100 gap-4">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-indigo-50 border border-indigo-200 text-indigo-600 rounded-xl">
              <Layers className="w-5 h-5 animate-pulse" />
            </div>
            <div>
              <h2 className="text-lg font-black text-zinc-900 tracking-tight flex items-center gap-2">
                SKU 库存批量快速查询系统
              </h2>
              <p className="text-xs text-zinc-500 mt-0.5">
                支持查询系统预载库存，以及导入外部多工作表 Excel 汇总，并包含截屏/图片 OCR 提取。
              </p>
            </div>
          </div>

          <div className="flex self-start sm:self-center bg-zinc-100 p-0.5 rounded-lg border border-zinc-200">
            <button
              onClick={() => {
                setQueryMode("system");
                setQueryResults([]);
                setIsSearched(false);
              }}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all ${
                queryMode === "system"
                  ? "bg-white text-zinc-900 shadow-sm"
                  : "text-zinc-500 hover:text-zinc-900"
              }`}
            >
              直接查询系统库存
            </button>
            <button
              onClick={() => {
                setQueryMode("custom");
                setQueryResults([]);
                setIsSearched(false);
              }}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all ${
                queryMode === "custom"
                  ? "bg-white text-zinc-900 shadow-sm"
                  : "text-zinc-500 hover:text-zinc-900"
              }`}
            >
              单独上传外部底表汇总
            </button>
          </div>
        </div>

        <div className="p-6 grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Controls Column */}
          <div className="space-y-6 lg:col-span-1">
            {/* Step 1: Data Source Panel */}
            {queryMode === "system" ? (
              <div className="bg-indigo-50/50 rounded-xl p-5 border border-indigo-100/80 space-y-3">
                <div className="flex items-center justify-between">
                  <h3 className="text-sm font-bold text-indigo-900 flex items-center gap-1.5">
                    <Database className="w-4 h-4 text-indigo-500" />
                    系统底表库存接入中
                  </h3>
                  <span className="text-[10px] bg-indigo-100 text-indigo-800 px-2 py-0.5 rounded-md font-bold">
                    在线热同步
                  </span>
                </div>
                <p className="text-xs text-indigo-700 leading-relaxed">
                  系统将实时穿透您在主看板导入的库存详情，进行全仓检索：包括可用库存及库龄数据。
                </p>
                {hasSystemData ? (
                  <div className="flex items-center gap-1 text-xs text-emerald-700 font-bold bg-emerald-50 border border-emerald-100 p-2.5 rounded-lg">
                    <CheckCircle2 className="w-4 h-4 text-emerald-500 flex-shrink-0" />
                    系统已正确加载 {Object.keys(systemInventoryMap).length} 个 SKU 仓库详情。
                  </div>
                ) : (
                  <div className="flex items-start gap-2.5 text-xs text-amber-800 bg-amber-50 border border-amber-100 p-3 rounded-lg leading-relaxed">
                    <AlertCircle className="w-4 h-4 text-amber-500 mt-0.5 flex-shrink-0" />
                    <div>
                      系统内暂未导入库存数据。您可以前往 <strong>“资费解析”</strong> 板块上传库存库，或切换上方
                      “<strong>单独上传外部底表汇总</strong>” 使用特定表格。
                    </div>
                  </div>
                )}
              </div>
            ) : (
              <div className="bg-zinc-50 p-5 rounded-xl border border-zinc-200 space-y-4">
                <h3 className="text-sm font-bold text-zinc-800 flex items-center gap-1.5">
                  <Upload className="w-4 h-4 text-zinc-500" />
                  第 1 步: 导入本地底表 (CSV / Excel)
                </h3>

                {/* Drop Zone */}
                <div
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                  onClick={() => fileInputRef.current?.click()}
                  className={`border-2 border-dashed rounded-xl p-6 text-center cursor-pointer transition-all ${
                    isDragOver
                      ? "border-indigo-500 bg-indigo-50/50"
                      : "border-zinc-300 bg-white hover:border-zinc-400"
                  }`}
                >
                  <input
                    type="file"
                    ref={fileInputRef}
                    onChange={handleFileChange}
                    accept=".csv, .xls, .xlsx"
                    className="hidden"
                  />
                  <Upload className="w-8 h-8 text-zinc-400 mx-auto mb-2" />
                  <p className="text-xs font-semibold text-zinc-700">拖拽文件到此处或点击上传</p>
                  <p className="text-[10px] text-zinc-400 mt-1">支持 CSV, XLS, XLSX 多子工作表</p>
                </div>

                {/* Multi-sheet Selector */}
                {sheetNames.length > 0 && (
                  <div className="p-3 bg-indigo-50/60 rounded-xl border border-indigo-100 space-y-1.5">
                    <label className="block text-xs font-bold text-indigo-900">
                      📄 切换分析子工作表 (Sheet)：
                    </label>
                    <select
                      value={currentSheet}
                      onChange={handleSheetChange}
                      className="w-full bg-white border border-zinc-200 rounded-lg p-2 text-xs font-medium text-zinc-800 focus:outline-none focus:ring-1 focus:ring-indigo-400"
                    >
                      {sheetNames.map((name) => (
                        <option key={name} value={name}>
                          {name}
                        </option>
                      ))}
                    </select>
                  </div>
                )}

                {/* Parse Status Indicator */}
                <div
                  className={`text-xs p-3 rounded-lg border flex items-start gap-2 ${
                    isSuccessStatus
                      ? "bg-emerald-50 border-emerald-200 text-emerald-800"
                      : isErrorStatus
                      ? "bg-rose-50 border-rose-200 text-rose-800"
                      : "bg-zinc-100 border-zinc-200 text-zinc-600"
                  }`}
                >
                  {isSuccessStatus ? (
                    <CheckCircle2 className="w-4 h-4 text-emerald-600 mt-0.5 flex-shrink-0" />
                  ) : isErrorStatus ? (
                    <AlertCircle className="w-4 h-4 text-rose-600 mt-0.5 flex-shrink-0" />
                  ) : (
                    <HelpCircle className="w-4 h-4 text-zinc-400 mt-0.5 flex-shrink-0" />
                  )}
                  <p className="font-semibold leading-relaxed break-all">{fileParseStatus}</p>
                </div>
              </div>
            )}

            {/* Custom Mode Selectors */}
            {queryMode === "custom" && customData.length > 0 && (
              <div className="bg-zinc-50 p-5 rounded-xl border border-zinc-200 space-y-4">
                <h3 className="text-sm font-bold text-zinc-800 flex items-center gap-1.5">
                  <Database className="w-4 h-4 text-zinc-500" />
                  第 2 步: 映射表头与求和字段
                </h3>

                <div className="space-y-3 font-medium text-zinc-700">
                  {/* SKU Column Map */}
                  <div className="space-y-1">
                    <span className="text-xs text-zinc-500">基础 SKU 匹配列 / 判定键：</span>
                    <select
                      value={skuColumn}
                      onChange={(e) => setSkuColumn(e.target.value)}
                      className="w-full bg-white border border-zinc-200 rounded-lg p-2 text-xs focus:ring-1 focus:ring-indigo-400"
                    >
                      {headers.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Summary Value Column Map */}
                  <div className="space-y-1">
                    <span className="text-xs text-zinc-500">累加计算列 (库存数量) ：</span>
                    <select
                      value={sumColumn}
                      onChange={(e) => setSumColumn(e.target.value)}
                      className="w-full bg-white border border-zinc-200 rounded-lg p-2 text-xs focus:ring-1 focus:ring-indigo-400"
                    >
                      {headers.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
            )}

            {/* Matching Rules Options */}
            <div className="bg-zinc-50 p-5 rounded-xl border border-zinc-200 space-y-3">
              <h3 className="text-xs font-extrabold text-zinc-400 uppercase tracking-widest">
                高级设定与匹配规则
              </h3>
              <div className="flex items-center justify-between border-b border-zinc-100 pb-2.5">
                <span className="text-xs font-semibold text-zinc-700">匹配判定模式</span>
                <div className="bg-zinc-200 p-0.5 rounded-lg flex gap-1">
                  <button
                    onClick={() => setMatchingMode("fuzzy")}
                    className={`px-2 py-1 text-[10px] font-bold rounded-md transition-all ${
                      matchingMode === "fuzzy"
                        ? "bg-white text-zinc-900 shadow-sm"
                        : "text-zinc-500 hover:text-zinc-900"
                    }`}
                  >
                    模糊包容
                  </button>
                  <button
                    onClick={() => setMatchingMode("exact")}
                    className={`px-2 py-1 text-[10px] font-bold rounded-md transition-all ${
                      matchingMode === "exact"
                        ? "bg-white text-zinc-900 shadow-sm"
                        : "text-zinc-500 hover:text-zinc-900"
                    }`}
                  >
                    完全精确
                  </button>
                </div>
              </div>
              <p className="text-[10px] text-zinc-400 leading-relaxed mt-1">
                {matchingMode === "fuzzy"
                  ? "💡 模糊模式下，只要表格内的SKU含有你输入的关键词（或反向包含），就会计入累加求和。"
                  : "💡 精确模式下，表格内单元格必须与你输入的SKU完全相同（不区分大小写，自动去空）。"}
              </p>
            </div>
          </div>

          {/* Textarea Input Column */}
          <div className="lg:col-span-2 space-y-6">
            <div className="bg-zinc-50 p-6 rounded-xl border border-zinc-200 flex flex-col h-full justify-between">
              <div>
                <div className="flex items-center justify-between mb-3">
                  <h3 className="text-sm font-bold text-zinc-800 flex items-center gap-1.5">
                    <Search className="w-4 h-4 text-zinc-500" />
                    需批量查询的 SKU
                  </h3>

                  {/* OCR trigger button */}
                  <div className="flex items-center">
                    <label
                      htmlFor="image-ocr-file"
                      className={`cursor-pointer inline-flex items-center gap-1 text-xs text-indigo-700 bg-indigo-50 border border-indigo-200 hover:bg-indigo-100 hover:scale-[1.02] active:scale-95 transition-all py-1 px-2.5 rounded-lg font-semibold shadow-sm ${
                        ocrLoading ? "opacity-60 cursor-not-allowed" : ""
                      }`}
                    >
                      {ocrLoading ? (
                        <div className="w-3.5 h-3.5 border-2 border-indigo-600 border-t-transparent rounded-full animate-spin"></div>
                      ) : (
                        <Camera className="w-3.5 h-3.5" />
                      )}
                      <span>{ocrLoading ? "识别中..." : "快照/截图批量文字提取"}</span>
                    </label>
                    <input
                      id="image-ocr-file"
                      ref={imageInputRef}
                      type="file"
                      accept="image/*"
                      onChange={handleImageOcrChange}
                      disabled={ocrLoading}
                      className="hidden"
                    />
                  </div>
                </div>

                <div className="relative">
                  <textarea
                    value={skuInput}
                    onChange={(e) => setSkuInput(e.target.value)}
                    rows={10}
                    placeholder="请输入或由上方OCR扫描追加基础 SKU，每行一个。例如：&#10;EB-C-1-HG9015&#10;SG4979&#10;AM-D-5-T328"
                    className="w-full bg-white border border-zinc-200 rounded-xl p-4 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-indigo-400 focus:border-transparent tracking-wide leading-relaxed shadow-inner"
                  />
                  {skuInput.trim().length > 0 && (
                    <button
                      onClick={() => setSkuInput("")}
                      className="absolute bottom-4 right-4 bg-zinc-100 hover:bg-zinc-200 text-zinc-500 font-bold p-1 rounded transition-colors text-xs flex items-center gap-1 shadow-sm border border-zinc-200"
                    >
                      <Trash2 className="w-3 h-3" /> 清空列表
                    </button>
                  )}
                </div>
              </div>

              <div className="mt-5 flex items-center justify-between gap-4">
                <div className="text-xs text-zinc-500 font-bold">
                  {skuInput.split("\n").filter((s) => s.trim().length > 0).length} 个待查项目
                </div>

                <button
                  onClick={handleQuery}
                  disabled={isSearching || (queryMode === "custom" && customData.length === 0)}
                  className={`flex-1 sm:flex-initial flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold py-3 px-8 rounded-xl shadow-md cursor-pointer transition-all hover:scale-[1.01] active:scale-95 text-sm ${
                    isSearching || (queryMode === "custom" && customData.length === 0)
                      ? "opacity-60 bg-zinc-400 cursor-not-allowed"
                      : "shadow-indigo-200"
                  }`}
                >
                  {isSearching ? (
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                  ) : (
                    <CheckSquare className="w-4 h-4" />
                  )}
                  {isSearching ? "正在检索对齐..." : "一键批量穿透查询"}
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Results Section */}
      {isSearched && (
        <div className="bg-white rounded-xl shadow-sm border border-zinc-200 overflow-hidden animate-in fade-in slide-in-from-bottom-4 duration-500">
          <div className="p-5 bg-zinc-50/50 border-b border-zinc-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-center gap-2">
              <CheckCircle2 className="w-5 h-5 text-emerald-600" />
              <h3 className="font-extrabold text-zinc-900 text-[15px]">
                检索分析结果汇总 ({queryResults.length} 项)
              </h3>
            </div>

            <div className="flex items-center gap-2 justify-end">
              <button
                onClick={handleCopyResults}
                className="flex items-center gap-1.5 text-xs text-zinc-700 bg-white hover:bg-zinc-50 border border-zinc-200 px-3 py-1.5 rounded-lg font-bold shadow-sm transition-all"
              >
                <Copy className="w-3.5 h-3.5" />
                复制表格格式
              </button>
              <button
                onClick={handleExportCSV}
                className="flex items-center gap-1.5 text-xs text-indigo-700 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 px-3 py-1.5 rounded-lg font-bold shadow-sm transition-all"
              >
                <FileDown className="w-3.5 h-3.5" />
                导出 CSV 结果文件
              </button>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-zinc-200 text-left border-collapse">
              <thead className="bg-zinc-50 font-bold text-xs text-zinc-600 tracking-wider">
                <tr>
                  <th scope="col" className="px-6 py-4 border-b border-zinc-100">
                    序号
                  </th>
                  <th scope="col" className="px-6 py-4 border-b border-zinc-100">
                    查询 SKU
                  </th>
                  <th scope="col" className="px-6 py-4 border-b border-zinc-100">
                    {queryMode === "system" ? "系统整合总库存" : `汇总求和 [${sumColumn}]`}
                  </th>
                  <th scope="col" className="px-6 py-4 border-b border-zinc-100">
                    匹配明细 / 对应记录仓库
                  </th>
                  <th scope="col" className="px-6 py-4 border-b border-zinc-100 text-center">
                    操作偏好
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-zinc-100 bg-white">
                {queryResults.map((res, index) => {
                  const hasStock = res.stockTotal > 0;
                  return (
                    <tr
                      key={index}
                      className="hover:bg-zinc-50/50 transition-colors text-sm font-medium"
                    >
                      <td className="px-6 py-4 text-zinc-400 font-mono text-xs">
                        {index + 1}
                      </td>
                      <td className="px-6 py-4 text-zinc-900 font-mono tracking-tight font-bold">
                        {res.querySku}
                      </td>
                      <td className={`px-6 py-4 font-mono font-extrabold ${hasStock ? "text-indigo-600 text-base" : "text-zinc-400"}`}>
                        {res.stockTotal}
                      </td>
                      <td className="px-6 py-4 text-xs text-zinc-500 leading-normal max-w-sm truncate" title={res.detailText}>
                        {res.matchStatus === "matched" ? (
                          <span className={`${hasStock ? "text-zinc-700" : "text-zinc-400"}`}>
                            {res.detailText}
                          </span>
                        ) : (
                          <span className="text-rose-500 bg-rose-50 border border-rose-100 px-2 py-0.5 rounded-md text-[10px] font-bold">
                            {res.detailText}
                          </span>
                        )}
                      </td>
                      <td className="px-6 py-4 text-center">
                        <button
                          onClick={() => {
                            if (onAutoSwitchToQuery) {
                              onAutoSwitchToQuery(res.querySku);
                            }
                          }}
                          className="text-[10px] font-extrabold bg-zinc-100 tracking-wide hover:bg-indigo-600 hover:text-white border border-zinc-200 hover:border-indigo-600 text-zinc-600 px-2.5 py-1 rounded-md shadow-sm transition-all inline-flex items-center gap-1 cursor-pointer"
                        >
                          <Sparkles className="w-3 h-3" /> 去资费计算
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
