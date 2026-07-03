import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useEffect } from "react";
import { useAppState } from "@/lib/app-context";
import { AppShell } from "@/components/xls-vcard/AppShell";
import { ExportStep } from "@/components/xls-vcard/XlsToVcardApp";

export const Route = createFileRoute("/export")({
  component: ExportRoute,
});

function ExportRoute() {
  const {
    parsed, filteredRows, fileName, setFileName,
    previewRow, cfg, headerMap, splitPhones,
    previewRowIdx, setPreviewRowIdx, download, reset,
  } = useAppState();
  const navigate = useNavigate();

  useEffect(() => {
    if (!parsed) navigate({ to: "/" });
  }, [parsed, navigate]);

  if (!parsed) return null;

  return (
    <AppShell>
      <ExportStep
        count={filteredRows.length}
        fileName={fileName}
        onFileNameChange={setFileName}
        onBack={() => navigate({ to: "/map" })}
        onDownload={download}
        onReset={() => { reset(); navigate({ to: "/" }); }}
        previewRow={previewRow}
        cfg={cfg}
        headerMap={headerMap}
        splitPhones={splitPhones}
        rowIdx={previewRowIdx}
        setRowIdx={setPreviewRowIdx}
        total={filteredRows.length}
      />
    </AppShell>
  );
}
