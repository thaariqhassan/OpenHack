import { promises as fs } from "fs";
import os from "os";
import path from "path";
import { NextRequest, NextResponse } from "next/server";
import { execFile } from "child_process";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const pdfEntries = formData.getAll("pdfFiles");
    const pdfFiles = pdfEntries.filter((entry): entry is File => entry instanceof File);

    if (pdfFiles.length === 0) {
      return NextResponse.json({ error: "Please upload at least one PDF." }, { status: 400 });
    }

    const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "plotvisionaries-"));
    const pdfDir = path.join(tempDir, "pdfs");
    await fs.mkdir(pdfDir, { recursive: true });

    for (const file of pdfFiles) {
      const safeName = file.name.replace(/[^a-zA-Z0-9._-]/g, "_");
      const destination = path.join(pdfDir, safeName);
      const buffer = Buffer.from(await file.arrayBuffer());
      await fs.writeFile(destination, buffer);
    }

    const outputPath = path.join(tempDir, "capture_coordinates.xlsx");
    const scriptPath = path.join(process.cwd(), "backend", "CaptureCoordinates.py");
    const pythonExecutable = path.join(
      process.cwd(),
      ".venv",
      process.platform === "win32" ? "Scripts/python.exe" : "bin/python"
    );

    await new Promise<void>((resolve, reject) => {
      execFile(
        pythonExecutable,
        [scriptPath, pdfDir, outputPath],
        { cwd: process.cwd() },
        (error, _stdout, stderr) => {
          if (error) {
            reject(new Error(stderr || error.message));
            return;
          }
          resolve();
        }
      );
    });

    const fileStat = await fs.stat(outputPath).catch(() => null);
    if (!fileStat || fileStat.size === 0) {
      return NextResponse.json(
        { error: "Capture failed: generated Excel file is missing or empty." },
        { status: 500 }
      );
    }

    const fileData = await fs.readFile(outputPath);
    return new NextResponse(fileData, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="capture_coordinates.xlsx"`,
        "Content-Length": fileStat.size.toString(),
      },
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Capture coordinates failed";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
