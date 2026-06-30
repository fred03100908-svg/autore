import { NextResponse } from "next/server";

const BACKEND_TIMEOUT_MS = 30_000;

function extractErrorMessage(data: unknown, fallback: string): string {
  if (!data || typeof data !== "object") {
    return fallback;
  }

  const objectData = data as Record<string, unknown>;
  const candidates = [objectData.message, objectData.detail, objectData.error];

  for (const candidate of candidates) {
    if (typeof candidate === "string" && candidate.trim()) {
      return candidate;
    }
  }

  return fallback;
}

async function readBackendError(response: Response, fallback: string): Promise<string> {
  const contentType = response.headers.get("content-type") || "";

  if (contentType.includes("application/json")) {
    const jsonData = await response.json().catch(() => null);
    return extractErrorMessage(jsonData, fallback);
  }

  const textData = await response.text().catch(() => "");
  return textData.trim() || fallback;
}

export async function POST(request: Request) {
  const formData = await request.formData();

  const excel = formData.get("excel");
  const template = formData.get("template");

  if (!(excel instanceof File) || !(template instanceof File)) {
    return NextResponse.json(
      {
        ok: false,
        message: "엑셀 파일과 한글 템플릿 파일을 모두 업로드해야 합니다.",
      },
      { status: 400 }
    );
  }

  const backendUrl = process.env.BACKEND_API_URL || "";

  if (!backendUrl) {
    return NextResponse.json(
      {
        ok: false,
        message:
          "BACKEND_API_URL 환경변수가 설정되지 않았습니다. 배포된 Python 백엔드 주소를 넣어야 합니다.",
      },
      { status: 503 }
    );
  }

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), BACKEND_TIMEOUT_MS);

  try {
    const response = await fetch(`${backendUrl}/parse-excel`, {
      method: "POST",
      headers: {
        "bypass-tunnel-reminder": "true",
      },
      body: formData,
      signal: controller.signal,
    });

    if (!response.ok) {
      const message = await readBackendError(
        response,
        "엑셀 분석 백엔드에 연결하지 못했습니다."
      );

      return NextResponse.json(
        {
          ok: false,
          message,
        },
        { status: response.status }
      );
    }

    const data = await response.json().catch(() => null);

    if (!data || typeof data !== "object") {
      return NextResponse.json(
        {
          ok: false,
          message: "백엔드 응답 형식이 올바르지 않습니다.",
        },
        { status: 502 }
      );
    }

    const typedData = data as Record<string, unknown>;

    return NextResponse.json(
      {
        ok: typedData.ok ?? true,
        message: typedData.message,
        excel_filename: typedData.excel_filename,
        template_filename: typedData.template_filename,
        available_sheets: typedData.available_sheets,
        selected_sheets: typedData.selected_sheets,
        basic_info: typedData.basic_info,
      },
      { status: response.status }
    );
  } catch (error) {
    if (error instanceof Error && error.name === "AbortError") {
      return NextResponse.json(
        {
          ok: false,
          message: "백엔드 응답 시간이 초과되었습니다. 잠시 후 다시 시도해 주세요.",
        },
        { status: 504 }
      );
    }

    return NextResponse.json(
      {
        ok: false,
        message: "백엔드 연결 중 네트워크 오류가 발생했습니다.",
      },
      { status: 502 }
    );
  } finally {
    clearTimeout(timeoutId);
  }
}