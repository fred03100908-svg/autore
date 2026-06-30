import { NextResponse } from "next/server";

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

  return NextResponse.json({
    ok: true,
    message:
      "업로드가 완료되었습니다. 다음 단계에서 실제 HWP/HWPX 보고서 생성 엔진을 연결하면 됩니다.",
    files: {
      excel: excel.name,
      template: template.name,
    },
  });
}