"use client";

import { useRef, useState } from "react";
import { motion, useScroll, useTransform } from "motion/react";
import {
  ArrowUpRight,
  CheckCircle2,
  FileSpreadsheet,
  FileText,
  Loader2,
  Sparkles,
  UploadCloud,
} from "lucide-react";

type GenerateStatus = "idle" | "uploading" | "complete" | "error";

export default function Home() {
  const generatorRef = useRef<HTMLElement | null>(null);

  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [hwpFile, setHwpFile] = useState<File | null>(null);
  const [status, setStatus] = useState<GenerateStatus>("idle");
  const [message, setMessage] = useState("");

  const { scrollYProgress } = useScroll();
  const heroMove = useTransform(scrollYProgress, [0, 0.35], [0, -120]);
  const blurMove = useTransform(scrollYProgress, [0, 0.35], [0, 90]);

  const canGenerate = !!excelFile && !!hwpFile && status !== "uploading";

  const scrollToGenerator = () => {
    generatorRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  const handleGenerate = async () => {
    if (!excelFile || !hwpFile) return;

    try {
      setStatus("uploading");
      setMessage("파일을 Python 백엔드로 보내는 중입니다.");

      const formData = new FormData();
      formData.append("excel", excelFile);
      formData.append("template", hwpFile);

      const res = await fetch("http://127.0.0.1:8000/parse-excel", {
        method: "POST",
        body: formData,
      });

      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.message || "보고서 생성에 실패했습니다.");
      }

      console.log("Backend response:", data);

      setStatus("complete");
      setMessage(data.message || "엑셀 파일을 성공적으로 읽었습니다.");
    } catch (error) {
      console.error(error);

      setStatus("error");
      setMessage(
        error instanceof Error
          ? error.message
          : "백엔드 연결 중 알 수 없는 오류가 발생했습니다."
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
            현재 단계는 프론트 화면에서 업로드한 엑셀과 한글 템플릿을
            Python/FastAPI 백엔드로 전송하고, 백엔드가 엑셀 파일을 읽는지
            확인하는 단계입니다.
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
                setStatus("idle");
                setMessage("");
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
                setStatus("idle");
                setMessage("");
              }}
            />
          </div>

          <div className="process-row">
            <ProcessStep number="01" title="Read Excel" active={!!excelFile} />
            <ProcessStep
              number="02"
              title="Map Values"
              active={!!excelFile && !!hwpFile}
            />
            <ProcessStep
              number="03"
              title="Generate HWP"
              active={status === "complete"}
            />
          </div>

          <div className="generate-area">
            <div className={`status-text ${status}`}>
              {status === "idle" &&
                "파일 2개를 업로드하면 생성 버튼이 활성화됩니다."}
              {status === "uploading" && message}
              {status === "complete" && message}
              {status === "error" && message}
            </div>

            <button
              className="generate-button"
              disabled={!canGenerate}
              onClick={handleGenerate}
            >
              {status === "uploading" ? (
                <>
                  <Loader2 className="spin" size={20} />
                  GENERATING
                </>
              ) : status === "complete" ? (
                <>
                  <CheckCircle2 size={20} />
                  COMPLETE
                </>
              ) : (
                <>
                  <Sparkles size={20} />
                  GENERATE REPORT
                </>
              )}
            </button>
          </div>
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
type ParsedData = {
  ok: boolean;
  message: string;
  basic_info?: Record<string, string | number | null>;
  selected_sheets?: {
    input1: string;
    reports: string[];
  };
  available_sheets?: string[];
};