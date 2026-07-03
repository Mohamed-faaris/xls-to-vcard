import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useEffect } from "react";
import { useAppState } from "@/lib/app-context";
import { AppShell } from "@/components/xls-vcard/AppShell";
import { MapStep } from "@/components/xls-vcard/XlsToVcardApp";

export const Route = createFileRoute("/map")({
  component: MapRoute,
});

function MapRoute() {
  const {
    parsed, cfg, dispatch, filters, setFilters, filteredRows,
    previewRow, previewRowIdx, setPreviewRowIdx, headerMap,
    splitPhones, setSplitPhones,
  } = useAppState();
  const navigate = useNavigate();

  useEffect(() => {
    if (!parsed) navigate({ to: "/" });
  }, [parsed, navigate]);

  if (!parsed) return null;

  return (
    <AppShell>
      <MapStep
        parsed={parsed}
        cfg={cfg}
        dispatch={dispatch}
        filters={filters}
        setFilters={setFilters}
        filteredRows={filteredRows}
        previewRow={previewRow}
        previewRowIdx={previewRowIdx}
        setPreviewRowIdx={setPreviewRowIdx}
        headerMap={headerMap}
        onBack={() => navigate({ to: "/preview" })}
        onNext={() => navigate({ to: "/export" })}
        splitPhones={splitPhones}
        setSplitPhones={setSplitPhones}
      />
    </AppShell>
  );
}
