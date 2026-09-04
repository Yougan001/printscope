'use client';

import { useCallback, useEffect, useRef, useState } from 'react';
import Link from 'next/link';
import {
  Printer,
  FileSpreadsheet,
  Download,
  ArrowUpRight,
  X,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select';
import { createSample } from '@/core/sample.mjs';
import { rangeLabel } from '@/core/ranges.mjs';
import type { Box, WorkbookReport, SheetReport } from '@/core/report';

function save(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function Footprint({ sheet }: { sheet: SheetReport }) {
  const ranges = sheet.ranges!;
  const boxes = [
    ...Object.values(ranges),
    ...(sheet.printArea?.ranges ?? []),
  ].filter((box): box is Box => !!box);
  const rows = Math.max(1, ...boxes.map((box) => box.bottom));
  const columns = Math.max(1, ...boxes.map((box) => box.right));
  const rect = (box: Box) => ({
    x: ((box.left - 1) / columns) * 760 + 20,
    y: ((box.top - 1) / rows) * 300 + 16,
    width: Math.max(1, ((box.right - box.left + 1) / columns) * 760),
    height: Math.max(1, ((box.bottom - box.top + 1) / rows) * 300),
  });
  return (
    <div className="footprint">
      <svg
        viewBox="0 0 800 344"
        // Inline SVG needs an image role to expose its coordinate description.
        // eslint-disable-next-line jsx-a11y/prefer-tag-over-role
        role="img"
        aria-label={`Cell-coordinate map, ending at row ${rows}, column ${columns}. Populated ${rangeLabel(ranges.populated) ?? 'empty'}, candidate-visible ${rangeLabel(ranges.candidates) ?? 'empty'}.`}
      >
        <defs>
          <pattern
            id="grid"
            width="38"
            height="20"
            patternUnits="userSpaceOnUse"
          >
            <path
              d="M 38 0 L 0 0 0 20"
              fill="none"
              stroke="#e7ebf5"
              strokeWidth="1"
            />
          </pattern>
        </defs>
        <rect
          x="20"
          y="16"
          width="760"
          height="300"
          fill="url(#grid)"
          stroke="#cbd3e6"
        />
        {sheet.printArea?.ranges.map((box, index) => (
          <rect
            key={index}
            {...rect(box)}
            fill="#fff4e3"
            fillOpacity=".6"
            stroke="#a45618"
            strokeWidth="2"
            strokeDasharray="7 5"
          />
        ))}
        {ranges.populated && (
          <rect
            {...rect(ranges.populated)}
            fill="none"
            stroke="#283143"
            strokeWidth="2"
          />
        )}
        {ranges.candidates && (
          <rect
            {...rect(ranges.candidates)}
            fill="#dce4ff"
            fillOpacity=".9"
            stroke="#253db3"
            strokeWidth="2"
          />
        )}
        {ranges.drawingBounds && (
          <rect
            {...rect(ranges.drawingBounds)}
            fill="none"
            stroke="#8c367b"
            strokeWidth="2"
            strokeDasharray="3 3"
          />
        )}
        <text x="20" y="337" fill="#626b80" fontSize="13">
          A1
        </text>
        <text x="780" y="337" textAnchor="end" fill="#626b80" fontSize="13">
          {columns} columns × {rows.toLocaleString()} rows
        </text>
      </svg>
      <div className="legend">
        <span>
          <i className="candidate" />
          Candidate-visible cells
        </span>
        <span>
          <i className="populated" />
          Populated cells
        </span>
        <span>
          <i className="print" />
          Print area
        </span>
        {ranges.drawingBounds && (
          <span>
            <i className="drawing" />
            Drawing anchors
          </span>
        )}
      </div>
      <p className="note">
        Coordinates, not physical page sizes. The background grid is schematic;
        small ranges may appear as a line.
      </p>
    </div>
  );
}

export default function Home() {
  const [report, setReport] = useState<WorkbookReport | null>(null);
  const [selected, setSelected] = useState(0);
  const [filename, setFilename] = useState('');
  const [sample, setSample] = useState(false);
  const [busy, setBusy] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const input = useRef<HTMLInputElement>(null);
  const active = useRef<Worker | null>(null);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const stop = useCallback(() => {
    active.current?.terminate();
    active.current = null;
    if (timer.current) clearTimeout(timer.current);
    timer.current = null;
  }, []);
  const inspect = useCallback(
    (file?: File) => {
      stop();
      setError('');
      setNotice('');
      setReport(null);
      setSelected(0);
      setFilename(file?.name ?? 'monthly-report.xlsx');
      setSample(!file);
      setBusy('Opening workbook…');
      try {
        const worker = new Worker(
          new URL('./inspect.worker.ts', import.meta.url),
          {
            type: 'module',
          },
        );
        active.current = worker;
        worker.onmessage = (
          event: MessageEvent<{
            report?: WorkbookReport;
            error?: string;
            status?: string;
          }>,
        ) => {
          if (active.current !== worker) return;
          if (event.data.status) {
            setBusy(event.data.status);
            return;
          }
          stop();
          setBusy('');
          if (event.data.error) setError(event.data.error);
          else if (event.data.report) setReport(event.data.report);
        };
        worker.onerror = () => {
          if (active.current !== worker) return;
          stop();
          setBusy('');
          setError(
            'The inspection worker could not finish. Reload the page or try a smaller workbook.',
          );
        };
        timer.current = setTimeout(() => {
          if (active.current !== worker) return;
          stop();
          setBusy('');
          setError(
            'Inspection stopped after 30 seconds. Try a smaller workbook.',
          );
        }, 30_000);
        if (file) worker.postMessage({ file });
        else {
          const bytes = createSample();
          worker.postMessage({ bytes }, [bytes.buffer]);
        }
      } catch {
        stop();
        setBusy('');
        setError('This browser could not start local workbook inspection.');
      }
    },
    [stop],
  );
  useEffect(() => {
    const initial = setTimeout(() => inspect(), 0);
    return () => {
      clearTimeout(initial);
      stop();
    };
  }, [inspect, stop]);
  const sheet = report?.sheets[selected];
  function clear() {
    stop();
    setBusy('');
    setReport(null);
    setFilename('');
    setError('');
    setNotice('Workbook cleared.');
    if (input.current) input.current.value = '';
  }
  function downloadReport() {
    if (!report) return;
    save(
      new Blob([JSON.stringify({ file: filename, ...report }, null, 2)], {
        type: 'application/json',
      }),
      'printscope-report.json',
    );
    setNotice(
      'Report download started. It includes sheet names and cell addresses, not cell contents.',
    );
  }
  function downloadSample() {
    const bytes = createSample();
    save(
      new Blob([bytes as Uint8Array<ArrayBuffer>], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
      'monthly-report.xlsx',
    );
    setNotice(
      'Sample workbook download started. Open it in Excel to inspect cell M85.',
    );
  }
  return (
    <main className="sheet-desk">
      <header>
        <Link className="wordmark" href="/">
          <Printer aria-hidden="true" />
          Printscope
        </Link>
        <a
          className="source-link"
          href="https://github.com/Yougan001/printscope"
        >
          Source <ArrowUpRight size={16} />
        </a>
      </header>
      <div className="intro">
        <div>
          <p className="eyebrow">EXCEL PRINT DIAGNOSTICS</p>
          <h1>Where did those extra pages come from?</h1>
          <p>
            Inspect an .xlsx file for distant cells, hidden-looking values and
            oversized print areas.
          </p>
        </div>
        <p className="local-note">
          Read-only.
          <br />
          Your workbook stays here.
        </p>
      </div>
      <div className="workspace">
        <aside aria-label="Workbook controls">
          <div className="import-box">
            <FileSpreadsheet size={28} aria-hidden="true" />
            <h2>Open a workbook</h2>
            <p>
              .xlsx · up to 20 MiB
              <br />
              No upload. No account.
            </p>
            <input
              ref={input}
              id="workbook-file"
              type="file"
              accept=".xlsx"
              className="sr-only"
              aria-label="Choose XLSX workbook"
              onChange={(event) => {
                const file = event.target.files?.[0];
                if (file) inspect(file);
                event.target.value = '';
              }}
            />
            <Button onClick={() => input.current?.click()} size="lg">
              Choose .xlsx
            </Button>
            <Button variant="link" onClick={() => inspect()}>
              Try the sample
            </Button>
          </div>
          {filename && (
            <div className="file-summary">
              <p className="eyebrow">
                {sample ? 'SAMPLE WORKBOOK' : 'SELECTED WORKBOOK'}
              </p>
              <p className="filename">{filename}</p>
              <Button variant="ghost" onClick={clear}>
                <X size={16} />
                Clear workbook
              </Button>
            </div>
          )}
          {report && (
            <>
              <label
                className="field-label"
                id="sheet-label"
                htmlFor="sheet-select"
              >
                Worksheet
              </label>
              <Select
                value={String(selected)}
                onValueChange={(value) =>
                  value !== null && setSelected(Number(value))
                }
              >
                <SelectTrigger
                  id="sheet-select"
                  aria-labelledby="sheet-label"
                  className="sheet-select"
                >
                  <SelectValue>{sheet?.name}</SelectValue>
                </SelectTrigger>
                <SelectContent>
                  {report.sheets.map((item, index) => (
                    <SelectItem key={index} value={String(index)}>
                      {item.name}
                      {item.state !== 'visible' ? ` (${item.state})` : ''}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
              <p className="note">
                {report.sheets.length} sheet(s) ·{' '}
                {report.storedCellCount.toLocaleString()} stored cells
              </p>
              <Button
                variant="outline"
                onClick={downloadReport}
                className="report-button"
              >
                <Download size={16} />
                Export JSON report
              </Button>
              <p className="note">
                Reports include filenames, sheet names and cell addresses.
                Review before sharing.
              </p>
            </>
          )}
          <details className="limits">
            <summary>File limits & privacy</summary>
            <p>
              50 sheets, 100,000 stored cells, 100,000 shared strings, 20,000
              styles. Oversized or damaged files are rejected.
            </p>
            <p>
              No workbooks are saved in browser storage or sent to a server.
              Clear the workbook or close this tab when finished. Hosting
              providers receive normal page requests.
            </p>
            <p>
              Unencrypted .xlsx only. No .xls, .xlsb or macro-enabled files.
            </p>
          </details>
        </aside>
        <section
          className="results"
          aria-label="Inspection results"
          aria-busy={!!busy}
        >
          {busy && (
            <div className="state-panel">
              <output className="busy-label">{busy}</output>
              <p>
                Checking the file locally. Large workbooks may take a moment.
              </p>
              <Button
                variant="outline"
                onClick={() => {
                  stop();
                  setBusy('');
                  setReport(null);
                  setNotice(
                    'Inspection cancelled. No partial report was kept.',
                  );
                }}
              >
                Cancel inspection
              </Button>
            </div>
          )}
          {error && (
            <div className="state-panel error" role="alert">
              <h2>Could not inspect this workbook</h2>
              <p>{error}</p>
              <p>
                No partial results are shown. Your original file is unchanged.
              </p>
            </div>
          )}
          {notice && <output className="status-note">{notice}</output>}
          {!busy && !error && !report && (
            <div className="state-panel">
              <h2>Choose a workbook to begin</h2>
              <p>
                Or try the sample: it contains a value in M85 hidden by the
                custom format <code>;;;</code>.
              </p>
            </div>
          )}
          {sheet && !sheet.supported && (
            <div className="state-panel">
              <h2>{sheet.name}</h2>
              <p>{sheet.reason}</p>
              <p>
                Select another worksheet. This sheet has not been diagnosed.
              </p>
            </div>
          )}
          {sheet?.supported && (
            <>
              <div className="map-heading">
                <div>
                  <p className="eyebrow">
                    {sample ? 'REAL SAMPLE FILE' : 'WORKBOOK INSPECTION'} /{' '}
                    {sheet.state}
                  </p>
                  <h2>{sheet.name}</h2>
                </div>
                <span>{sheet.findings?.length ?? 0} findings to review</span>
              </div>
              <div className="range-strip">
                <div>
                  <span>Populated cells</span>
                  <strong>
                    {rangeLabel(sheet.ranges?.populated) ?? 'Empty'}
                  </strong>
                </div>
                <div>
                  <span>Candidate-visible</span>
                  <strong>
                    {rangeLabel(sheet.ranges?.candidates) ?? 'Empty'}
                  </strong>
                </div>
                <div>
                  <span>Declared used range</span>
                  <strong>
                    {rangeLabel(sheet.ranges?.dimension) ?? 'Not stored'}
                  </strong>
                </div>
              </div>
              <Footprint sheet={sheet} />
              <div className="print-area">
                <span>Print area</span>
                <code>
                  {sheet.printArea?.expression ?? 'Not explicitly set'}
                </code>
                {!sheet.printArea?.resolved && (
                  <strong>Unresolved expression</strong>
                )}
              </div>
              <p className="note">
                Candidate-visible bounds exclude a few known hiding settings.
                They are not a recommended print area.
              </p>
              {sheet.findings?.map((finding, index) => (
                <article
                  className={`finding ${finding.level}`}
                  key={finding.id}
                >
                  <span>{String(index + 1).padStart(2, '0')}</span>
                  <div>
                    <p className="finding-label">
                      {finding.level === 'info' ? 'CONTEXT' : 'REVIEW IN EXCEL'}
                    </p>
                    <h3>{finding.title}</h3>
                    <p>{finding.evidence}</p>
                    <p className="action">
                      <strong>Next check</strong>
                      {finding.action}
                    </p>
                  </div>
                </article>
              ))}
              {!sheet.findings?.length && (
                <div className="state-panel">
                  <h3>No tracked clues found</h3>
                  <p>
                    This does not guarantee a particular page count. Check
                    Excel’s print preview and the limits below.
                  </p>
                </div>
              )}
              <details className="details" open={!!sheet.warnings?.length}>
                <summary>Print settings & inspection limits</summary>
                <dl className="settings">
                  <dt>Orientation</dt>
                  <dd>{sheet.pageSetup?.orientation}</dd>
                  <dt>Scaling</dt>
                  <dd>
                    {sheet.pageSetup?.fitToPage
                      ? `Fit mode · width ${sheet.pageSetup.fitToWidth ?? 'default'}, height ${sheet.pageSetup.fitToHeight ?? 'default'}`
                      : `${sheet.pageSetup?.scale ?? '100'}% (stored/default)`}
                  </dd>
                  <dt>Error printing</dt>
                  <dd>{sheet.pageSetup?.errors}</dd>
                  <dt>Manual row breaks</dt>
                  <dd>
                    {sheet.manualBreaks?.rows.join(', ') || 'None'}
                    {(sheet.manualBreaks?.rowCount ?? 0) > 100 ? ' …' : ''}
                  </dd>
                  <dt>Manual column breaks</dt>
                  <dd>
                    {sheet.manualBreaks?.columns.join(', ') || 'None'}
                    {(sheet.manualBreaks?.columnCount ?? 0) > 100 ? ' …' : ''}
                  </dd>
                </dl>
                <ul>
                  {[
                    ...(sheet.warnings ?? []),
                    ...(report?.limitations ?? []),
                  ].map((item, index) => (
                    <li key={index}>{item}</li>
                  ))}
                </ul>
              </details>
              {sample && (
                <div className="sample-download">
                  <p>Want to check the diagnosis yourself?</p>
                  <Button variant="outline" onClick={downloadSample}>
                    <Download size={16} />
                    Download sample .xlsx
                  </Button>
                </div>
              )}
            </>
          )}
        </section>
      </div>
      <footer>
        <p>
          An explanation, not a blind repair. Confirm the result in Excel before
          printing.
        </p>
        <a href="https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/blank-pages-unexpectedly-printed">
          Microsoft’s blank-page troubleshooting guide ↗
        </a>
      </footer>
    </main>
  );
}
