"use client";

import { useRef, useState } from "react";
import { motion, useScroll, useTransform } from "motion/react";
import {
  ArrowUpRight,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  FileText,
  Loader2,
  Sparkles,
  UploadCloud,
} from "lucide-react";

type GenerateStatus = "idle" | "uploading" | "complete" | "error";
type DownloadStatus = "idle" | "generating" | "complete" | "error";

type BasicValue = string | number | boolean | null;

type ParsedData = {
  ok: boolean;
  message: string;
  excel_filename?: string;
  template_filename?: string;
  available_sheets?: string[];
  selected_sheets?: {
    input1?: string;
    reports?: string[];
  };
  basic_info?: Record<string, BasicValue>;
};

export default function Home() {
  const generatorRef = useRef<HTMLElement | null>(null);

  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [hwpFile, setHwpFile] = useState<File | null>(null);
  const [status, setStatus] = useState<GenerateStatus>("idle");
  const [downloadStatus, setDownloadStatus] = useState<DownloadStatus>("idle");
  const [message, setMessage] = useState("");
  const [downloadMessage, setDownloadMessage] = useState("");
  const [parsedData, setParsedData] = useState<ParsedData | null>(null);

  const { scrollYProgress } = useScroll();
  const heroMove = useTransform(scrollYProgress, [0, 0.35], [0, -120]);
  const blurMove = useTransform(scrollYProgress, [0, 0.35], [0, 90]);

  const canAnalyze = !!excelFile && !!hwpFile && status !== "uploading";
  const canGenerate =
    !!excelFile &&
    !!hwpFile &&
    !!parsedData &&
    downloadStatus !== "generating";

  const scrollToGenerator = () => {
    generatorRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  const resetResult = () => {
    setStatus("idle");
    setDownloadStatus("idle");
    setMessage("");
    setDownloadMessage("");
    setParsedData(null);
  };

  const buildFormData = () => {
    if (!excelFile || !hwpFile) {
      throw new Error("엑셀 파일과 한글 템플릿 파일을 모두 업로드해야 합니다.");
    }

    const formData = new FormData();
    formData.append("excel", excelFile);
    formData.append("template", hwpFile);

    return formData;
  };

  const handleAnalyze = async () => {
    try {
      setStatus("uploading");
      setMessage("파일을 Python 백엔드로 보내는 중입니다.");
      setParsedData(null);

      const apiUrl = process.env.NEXT_PUBLIC_API_URL || "http://127.0.0.1:8000";
      const res = await fetch(`${apiUrl}/parse-excel`, {
        method: "POST",
        body: buildFormData(),
      });

      const data: ParsedData = await res.json();

      if (!res.ok || !data.ok) {
        throw new Error(data.message || "엑셀 분석에 실패했습니다.");
      }

      console.log("Backend response:", data);

      setParsedData(data);
      setStatus("complete");
      setMessage(data.message || "엑셀 파일을 성공적으로 읽었습니다.");
    } catch (error) {
      console.error(error);

      setStatus("error");
      setParsedData(null);
      setMessage(
        error instanceof Error
          ? error.message
          : "백엔드 연결 중 알 수 없는 오류가 발생했습니다."
      );
    }
  };

  const handleGenerateHwp = async () => {
    try {
      setDownloadStatus("generating");
      setDownloadMessage("HWP 보고서를 생성하는 중입니다.");

      const apiUrl = process.env.NEXT_PUBLIC_API_URL || "http://127.0.0.1:8000";
      const res = await fetch(`${apiUrl}/generate-report`, {
        method: "POST",
        body: buildFormData(),
      });

      if (!res.ok) {
        const errorData = await res.json();
        throw new Error(errorData.message || "HWP 보고서 생성에 실패했습니다.");
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);

      const contentDisposition = res.headers.get("content-disposition");
      let filename = "자동생성보고서.hwp";

      if (contentDisposition) {
        const match = contentDisposition.match(/filename="?([^"]+)"?/);
        if (match?.[1]) {
          filename = decodeURIComponent(match[1]);
        }
      }

      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();

      link.remove();
      window.URL.revokeObjectURL(url);

      setDownloadStatus("complete");
      setDownloadMessage("HWP 보고서 다운로드가 완료되었습니다.");
    } catch (error) {
      console.error(error);

      setDownloadStatus("error");
      setDownloadMessage(
        error instanceof Error
          ? error.message
          : "보고서 생성 중 알 수 없는 오류가 발생했습니다."
      );
    }
  };

  return (
    <main>
      <section className="hero">
        <div className="grain" />
        <motion.div className="hero-blob" style={{ y: blurMove }} />

        <nav className="top-nav">
          <span>©2026 / AUTOREPORT</span>
          <span>[ EXCEL TO HWP SYSTEM ]</span>
          <button onClick={scrollToGenerator}>[ CREATE HERE ]</button>
        </nav>

        <motion.div
          className="hero-content"
          style={{ y: heroMove }}
          initial={{ opacity: 0, y: 36 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.9, ease: [0.16, 1, 0.3, 1] }}
        >
          <div className="hero-left">
            <p className="hero-kicker">AUTO DOCUMENT GENERATOR</p>

            <h1 className="hero-title">
              2026
              <br />
              REPORT
              <br />
              MAKE HERE.
            </h1>
          </div>

          <div className="hero-right">
            <div className="made-card">
              <p>
                MADE WITH
                <br />
                <span>DATA</span>
              </p>
            </div>

            <button className="hero-cta" onClick={scrollToGenerator}>
              START GENERATE
              <ArrowUpRight size={22} />
            </button>
          </div>
        </motion.div>

        <div className="bottom-line">
          <span>SCROLL TO UPLOAD</span>
          <span>EXCEL / HWP / REPORT</span>
        </div>
      </section>

      <section className="generator-section" ref={generatorRef}>
        <div className="generator-bg-text">UPLOAD</div>

        <motion.div
          className="generator-header"
          initial={{ opacity: 0, y: 42 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true, amount: 0.4 }}
          transition={{ duration: 0.8, ease: [0.16, 1, 0.3, 1] }}
        >
          <p className="section-label">[ REPORT GENERATOR ]</p>
          <h2>
            엑셀 파일과 한글 템플릿을 넣으면
            <br />
            보고서 파일을 자동 생성합니다.
          </h2>
          <p className="section-desc">
            현재 단계는 엑셀 데이터를 확인한 뒤, HWP 템플릿 안의 치환 태그를
            실제 값으로 바꿔 자동 보고서를 다운로드하는 단계입니다.
          </p>
        </motion.div>

        <motion.div
          className="upload-panel"
          initial={{ opacity: 0, y: 52 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true, amount: 0.25 }}
          transition={{
            duration: 0.85,
            delay: 0.1,
            ease: [0.16, 1, 0.3, 1],
          }}
        >
          <div className="upload-grid">
            <UploadBox
              title="Excel File"
              description=".xlsx / .xlsm / .xls 파일을 업로드합니다."
              accept=".xlsx,.xlsm,.xls"
              file={excelFile}
              icon={<FileSpreadsheet size={28} />}
              onFileChange={(file) => {
                setExcelFile(file);
                resetResult();
              }}
            />

            <UploadBox
              title="HWP Template"
              description=".hwp / .hwpx 한글 템플릿을 업로드합니다."
              accept=".hwp,.hwpx"
              file={hwpFile}
              icon={<FileText size={28} />}
              onFileChange={(file) => {
                setHwpFile(file);
                resetResult();
              }}
            />
          </div>

          <div className="process-row">
            <ProcessStep number="01" title="Read Excel" active={!!excelFile} />
            <ProcessStep
              number="02"
              title="Check Data"
              active={!!parsedData}
            />
            <ProcessStep
              number="03"
              title="Generate HWP"
              active={downloadStatus === "complete"}
            />
          </div>

          <div className="generate-area">
            <div className={`status-text ${status}`}>
              {status === "idle" &&
                "파일 2개를 업로드하면 분석 버튼이 활성화됩니다."}
              {status === "uploading" && message}
              {status === "complete" && message}
              {status === "error" && message}
            </div>

            <button
              className="generate-button"
              disabled={!canAnalyze}
              onClick={handleAnalyze}
            >
              {status === "uploading" ? (
                <>
                  <Loader2 className="spin" size={20} />
                  ANALYZING
                </>
              ) : status === "complete" ? (
                <>
                  <CheckCircle2 size={20} />
                  DATA READY
                </>
              ) : (
                <>
                  <Sparkles size={20} />
                  ANALYZE EXCEL
                </>
              )}
            </button>
          </div>

          {parsedData && (
            <>
              <BasicInfoPreview data={parsedData} />

              <div className="download-panel">
                <div>
                  <p className="preview-label">[ HWP OUTPUT ]</p>
                  <h3>검토한 데이터로 HWP 보고서를 생성합니다.</h3>
                  <p className={`download-message ${downloadStatus}`}>
                    {downloadMessage ||
                      "템플릿에 {{학교명}}, {{교장}}, {{소재지}} 같은 태그가 있어야 값이 치환됩니다."}
                  </p>
                </div>

                <button
                  className="generate-button"
                  disabled={!canGenerate}
                  onClick={handleGenerateHwp}
                >
                  {downloadStatus === "generating" ? (
                    <>
                      <Loader2 className="spin" size={20} />
                      GENERATING HWP
                    </>
                  ) : downloadStatus === "complete" ? (
                    <>
                      <CheckCircle2 size={20} />
                      DOWNLOADED
                    </>
                  ) : (
                    <>
                      <Download size={20} />
                      GENERATE HWP
                    </>
                  )}
                </button>
              </div>
            </>
          )}
        </motion.div>
      </section>
    </main>
  );
}

function UploadBox({
  title,
  description,
  accept,
  file,
  icon,
  onFileChange,
}: {
  title: string;
  description: string;
  accept: string;
  file: File | null;
  icon: React.ReactNode;
  onFileChange: (file: File | null) => void;
}) {
  const [dragging, setDragging] = useState(false);
  const inputId = title.replaceAll(" ", "-").toLowerCase();

  return (
    <label
      htmlFor={inputId}
      className={`upload-box ${dragging ? "dragging" : ""} ${
        file ? "has-file" : ""
      }`}
      onDragOver={(e) => {
        e.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);

        const droppedFile = e.dataTransfer.files?.[0];

        if (droppedFile) {
          onFileChange(droppedFile);
        }
      }}
    >
      <input
        id={inputId}
        type="file"
        accept={accept}
        onChange={(e) => {
          const selectedFile = e.target.files?.[0] || null;
          onFileChange(selectedFile);
        }}
      />

      <div className="upload-icon">
        {file ? <CheckCircle2 size={28} /> : icon}
      </div>

      <div>
        <h3>{title}</h3>
        <p>{file ? file.name : description}</p>
      </div>

      <div className="upload-action">
        <UploadCloud size={18} />
        <span>{file ? "CHANGE FILE" : "DROP OR SELECT"}</span>
      </div>
    </label>
  );
}

function ProcessStep({
  number,
  title,
  active,
}: {
  number: string;
  title: string;
  active: boolean;
}) {
  return (
    <div className={`process-step ${active ? "active" : ""}`}>
      <span>{number}</span>
      <p>{title}</p>
    </div>
  );
}

function BasicInfoPreview({ data }: { data: ParsedData }) {
  const basicInfo = data.basic_info || {};
  const entries = Object.entries(basicInfo);

  return (
    <div className="preview-panel">
      <div className="preview-header">
        <div>
          <p className="preview-label">[ EXTRACTED DATA ]</p>
          <h3>엑셀에서 읽은 기본정보</h3>
        </div>

        <div className="sheet-badge">
          {data.selected_sheets?.input1 || "입력 시트 정보 없음"}
        </div>
      </div>

      <div className="file-summary">
        <div>
          <span>Excel</span>
          <p>{data.excel_filename || "-"}</p>
        </div>
        <div>
          <span>Template</span>
          <p>{data.template_filename || "-"}</p>
        </div>
      </div>

      <div className="info-grid">
        {entries.map(([key, value]) => (
          <div className="info-item" key={key}>
            <span>{key}</span>
            <strong>
              {value === null || value === "" || value === undefined
                ? "-"
                : String(value)}
            </strong>
          </div>
        ))}
      </div>

      <div className="report-sheet-list">
        <span>읽은 보고서 시트</span>
        <p>{data.selected_sheets?.reports?.join(", ") || "-"}</p>
      </div>

      <div className="report-sheet-list">
        <span>전체 시트 개수</span>
        <p>{data.available_sheets?.length || 0}개</p>
      </div>
    </div>
  );
}
