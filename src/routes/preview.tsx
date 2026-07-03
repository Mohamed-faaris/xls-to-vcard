import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useEffect } from "react";
import { useAppState } from "@/lib/app-context";
import { AppShell } from "@/components/xls-vcard/AppShell";
import { PreviewStep } from "@/components/xls-vcard/XlsToVcardApp";

export const Route = createFileRoute("/preview")({
  component: PreviewRoute,
});

function PreviewRoute() {
  const {
    wb, sheetName, setSheetName, skipRows, setSkipRows,
    firstRowIsHeader, setFirstRowIsHeader, parsed,
    filters, setFilters, filteredRows, enterMapping,
  } = useAppState();
  const navigate = useNavigate();

  useEffect(() => {
    if (!wb && !parsed) navigate({ to: "/" });
  }, [wb, parsed, navigate]);

  if (!wb && !parsed) return null;

  return (
    <AppShell>
      <PreviewStep
        wb={wb}
        sheetName={sheetName}
        setSheetName={setSheetName}
        skipRows={skipRows}
        setSkipRows={setSkipRows}
        firstRowIsHeader={firstRowIsHeader}
        setFirstRowIsHeader={setFirstRowIsHeader}
        parsed={parsed}
        filters={filters}
        setFilters={setFilters}
        filteredRows={filteredRows}
        onBack={() => navigate({ to: "/" })}
        onNext={() => { enterMapping(); navigate({ to: "/map" }); }}
      />
    </AppShell>
  );
}
