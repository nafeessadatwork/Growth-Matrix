/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useLayoutEffect, useId, useCallback, useRef, type Attributes, type ReactNode } from "react";
import { Download, RefreshCcw, Sun, Moon, ChevronLeft, ChevronRight, Calendar, Eye, Printer, X } from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import { saveAs } from "file-saver";
import { renderAsync } from "docx-preview";

import {
  AppraisalData,
  INITIAL_FACTORS,
  INITIAL_GOALS,
} from "./types";
import { cn, calculateAppraisal, formatScore } from "./lib/utils";

const STORAGE_KEY = "mindx_growth_matrix_data";
const THEME_KEY = "mindx_theme";

const DEFAULT_DATA: AppraisalData = {
  employee: {
    name: "",
    position: "",
    department: "",
    projectsManaged: "",
    appraisalType: "",
    appraisalPeriod: "",
  },
  reviewer: {
    name: "",
    position: "",
    department: "",
    appraisalDue: "",
  },
  factors: INITIAL_FACTORS,
  goals: INITIAL_GOALS,
  comments: {
    employee: "",
    reviewer: "",
    ceo: "",
  },
  signatures: {
    employee: { name: "", date: new Date().toISOString().split("T")[0] },
    reviewer: { name: "", date: new Date().toISOString().split("T")[0] },
    ceo: { name: "", date: new Date().toISOString().split("T")[0] },
  },
};

type Theme = "dark" | "light";

const TABS = ["info", "factors", "goals", "summary"] as const;

const DOCX_PREVIEW_OPTS = {
  inWrapper: true,
  hideWrapperOnPrint: true,
} as const;

export default function App() {
  const [data, setData] = useState<AppraisalData>(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved) as AppraisalData & {
          reviewer?: { projectsManaged?: string };
          employee?: { projectsManaged?: string };
        };
        // Back-compat: older drafts stored projectsManaged under reviewer.
        if (!parsed.employee?.projectsManaged && parsed.reviewer?.projectsManaged) {
          parsed.employee = { ...parsed.employee, projectsManaged: parsed.reviewer.projectsManaged };
        }
        if (parsed.reviewer && "projectsManaged" in parsed.reviewer) {
          delete (parsed.reviewer as { projectsManaged?: unknown }).projectsManaged;
        }
        // Back-compat: older drafts defaulted reviewer.department to "Management".
        // Clear it only when the reviewer section is otherwise blank (treat as an old default).
        if (
          parsed.reviewer?.department === "Management" &&
          !parsed.reviewer?.name &&
          !parsed.reviewer?.position &&
          !parsed.reviewer?.appraisalDue
        ) {
          parsed.reviewer = { ...parsed.reviewer, department: "" };
        }
        return parsed as AppraisalData;
      } catch (e) {
        return DEFAULT_DATA;
      }
    }
    return DEFAULT_DATA;
  });

  const [activeTab, setActiveTab] = useState<(typeof TABS)[number]>("info");
  const [isResetting, setIsResetting] = useState(false);
  const [theme, setTheme] = useState<Theme>(() => {
    const stored = localStorage.getItem(THEME_KEY);
    return stored === "light" ? "light" : "dark";
  });

  const [factorSlideIndex, setFactorSlideIndex] = useState(0);
  const [factorSlideDir, setFactorSlideDir] = useState(1);
  const [previewOpen, setPreviewOpen] = useState(false);
  const previewTitleId = useId();
  const modalDocxRef = useRef<HTMLDivElement>(null);
  const printSinkRef = useRef<HTMLDivElement>(null);
  const [previewDocLoading, setPreviewDocLoading] = useState(false);
  const [previewDocError, setPreviewDocError] = useState<string | null>(null);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  }, [data]);

  useEffect(() => {
    setFactorSlideIndex((i) => {
      const max = Math.max(0, data.factors.length - 1);
      if (i > max) return max;
      if (i < 0) return 0;
      return i;
    });
  }, [data.factors.length]);

  useLayoutEffect(() => {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem(THEME_KEY, theme);
  }, [theme]);

  useEffect(() => {
    if (!previewOpen) return;
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Escape") setPreviewOpen(false);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [previewOpen]);

  useEffect(() => {
    if (previewOpen) {
      const prev = document.body.style.overflow;
      document.body.style.overflow = "hidden";
      return () => {
        document.body.style.overflow = prev;
      };
    }
  }, [previewOpen]);

  const stats = calculateAppraisal(data);

  useEffect(() => {
    if (!previewOpen) return;
    const el = modalDocxRef.current;
    if (!el) return;
    let cancelled = false;
    setPreviewDocLoading(true);
    setPreviewDocError(null);
    el.innerHTML = "";
    (async () => {
      try {
        const { generateAppraisalDoc } = await import("./lib/docxGenerator");
        const blob = await generateAppraisalDoc(data);
        if (cancelled || !modalDocxRef.current) return;
        await renderAsync(blob, modalDocxRef.current, undefined, DOCX_PREVIEW_OPTS);
      } catch {
        if (!cancelled) setPreviewDocError("Could not render the same document as Download. Try again or use Download.");
      } finally {
        if (!cancelled) setPreviewDocLoading(false);
      }
    })();
    return () => {
      cancelled = true;
      if (modalDocxRef.current) modalDocxRef.current.innerHTML = "";
    };
  }, [previewOpen, data]);

  const handlePrint = useCallback(async () => {
    if (previewOpen) {
      window.print();
      return;
    }
    const el = printSinkRef.current;
    if (!el) {
      window.print();
      return;
    }
    try {
      el.innerHTML = "";
      const { generateAppraisalDoc } = await import("./lib/docxGenerator");
      const blob = await generateAppraisalDoc(data);
      await renderAsync(blob, el, undefined, DOCX_PREVIEW_OPTS);
      await new Promise<void>((resolve) => {
        requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
      });
      window.print();
    } catch {
      alert("Could not prepare print from the official DOCX. Use Download or open Preview first.");
    }
  }, [data, previewOpen]);

  const handleDownload = async () => {
    try {
      const { generateAppraisalDoc } = await import("./lib/docxGenerator");
      const blob = await generateAppraisalDoc(data);
      saveAs(blob, `Appraisal_${data.employee.name || "Form"}.docx`);
    } catch (error) {
      console.error("Failed to generate DOCX", error);
      alert("Failed to generate document. Please check console for details.");
    }
  };

  const resetForm = () => {
    if (!isResetting) {
      setIsResetting(true);
      setTimeout(() => setIsResetting(false), 3000);
      return;
    }

    localStorage.removeItem(STORAGE_KEY);
    setData(DEFAULT_DATA);
    setActiveTab("info");
    setIsResetting(false);
    window.location.reload();
  };

  const tabButtonClass = (tab: (typeof TABS)[number]) =>
    cn(
      "px-5 py-2 text-xs font-bold uppercase tracking-widest rounded-full transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page",
      activeTab === tab
        ? "bg-tab-active text-foreground shadow-sm"
        : "text-tab-inactive hover:text-foreground",
    );

  return (
    <div
      className={cn(
        "min-h-screen bg-page text-foreground font-sans relative",
        previewOpen && "print-preview-active",
      )}
    >
      <div className="print:hidden">
        <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-accent-warm via-accent to-accent-deep z-[60]" />

        <header className="bg-header backdrop-blur-md border-b border-border sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center min-h-20 py-3 gap-3">
            <div className="flex items-center gap-2 min-w-0 shrink">
              <div className="h-14 w-14 shrink-0 overflow-hidden flex items-center justify-center">
                <img
                  src={new URL("./assets/mindx-logo-white.png", import.meta.url).href}
                  alt="MindX"
                  className="h-14 w-14 object-contain drop-shadow-[0_2px_6px_rgba(0,0,0,0.45)]"
                  loading="eager"
                />
              </div>
              <div className="min-w-0 flex flex-col justify-center">
                <h1 className="text-sm sm:text-xl font-light tracking-tight text-foreground leading-none whitespace-nowrap truncate">
                  MINDX <span className="font-bold text-accent">Growth Matrix</span>
                </h1>
              </div>
            </div>

            <div className="hidden lg:flex bg-tab-pill p-1 rounded-full border border-border shrink-0">
              {TABS.map((tab) => (
                <button
                  key={tab}
                  type="button"
                  onClick={() => setActiveTab(tab)}
                  className={tabButtonClass(tab)}
                >
                  {tab}
                </button>
              ))}
            </div>

            <div className="flex items-center gap-2 sm:gap-3 shrink-0">
              <button
                type="button"
                onClick={() => setTheme((t) => (t === "dark" ? "light" : "dark"))}
                className="flex items-center gap-2 rounded-full border border-border bg-elevated px-3 py-2 text-[10px] font-bold uppercase tracking-widest text-muted hover:text-foreground hover:bg-surface transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page"
                title={theme === "dark" ? "Switch to Paper mode" : "Switch to Dark mode"}
              >
                {theme === "dark" ? <Sun className="h-4 w-4 text-accent-warm" /> : <Moon className="h-4 w-4 text-accent" />}
                <span className="hidden sm:inline">{theme === "dark" ? "Paper" : "Dark"}</span>
              </button>
              <button
                type="button"
                onClick={resetForm}
                className={cn(
                  "flex items-center gap-2 p-2 rounded-lg transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page",
                  isResetting
                    ? "bg-red-500/15 text-red-600 px-2 sm:px-3 border border-red-500/35"
                    : "text-subtle hover:text-red-600",
                )}
                title={isResetting ? "Click again to confirm reset" : "Reset Form"}
              >
                <RefreshCcw className={cn("w-5 h-5", isResetting && "animate-spin")} />
                {isResetting && (
                  <span className="text-[10px] font-black uppercase tracking-widest whitespace-nowrap">Confirm?</span>
                )}
              </button>
              <button
                type="button"
                onClick={() => setPreviewOpen(true)}
                className="flex items-center gap-2 rounded-full border border-border bg-elevated px-3 sm:px-4 py-2.5 font-bold text-xs uppercase tracking-widest text-muted hover:text-foreground hover:bg-surface transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page"
                aria-label="Preview printable appraisal"
              >
                <Eye className="w-4 h-4 shrink-0" />
                <span className="hidden sm:inline">Preview</span>
              </button>
              <button
                type="button"
                onClick={handlePrint}
                className="flex items-center gap-2 rounded-full border border-border bg-elevated px-3 sm:px-4 py-2.5 font-bold text-xs uppercase tracking-widest text-muted hover:text-foreground hover:bg-surface transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page"
                aria-label="Print appraisal"
              >
                <Printer className="w-4 h-4 shrink-0" />
                <span className="hidden sm:inline">Print</span>
              </button>
              <button
                type="button"
                onClick={handleDownload}
                className="flex items-center gap-2 rounded-full bg-accent px-4 sm:px-5 py-2.5 font-bold text-xs uppercase tracking-widest text-white shadow-md shadow-accent/30 hover:brightness-110 active:scale-[0.98] transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page"
              >
                <Download className="w-4 h-4" />
                <span className="hidden sm:inline">Download</span>
              </button>
            </div>
          </div>

          <div className="lg:hidden border-t border-border py-2 -mx-4 px-4 flex gap-2 overflow-x-auto">
            {TABS.map((tab) => (
              <button
                key={tab}
                type="button"
                onClick={() => setActiveTab(tab)}
                className={cn(tabButtonClass(tab), "shrink-0")}
              >
                {tab}
              </button>
            ))}
          </div>
        </div>
        </header>

        <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <AnimatePresence mode="wait">
          {activeTab === "info" && (
            <motion.div
              key="info"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-8"
            >
              <Section title="Basic Information">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                  <div className="space-y-6">
                    <h3 className="text-sm font-semibold uppercase tracking-wider text-muted">Employee Details</h3>
                    <InputField
                      label="Full Name"
                      value={data.employee.name}
                      onChange={(val) => setData({ ...data, employee: { ...data.employee, name: val } })}
                      placeholder="e.g. John Doe"
                    />
                    <InputField
                      label="Position"
                      value={data.employee.position}
                      onChange={(val) => setData({ ...data, employee: { ...data.employee, position: val } })}
                      placeholder="e.g. Senior Developer"
                    />
                    <InputField
                      label="Department"
                      value={data.employee.department}
                      onChange={(val) => setData({ ...data, employee: { ...data.employee, department: val } })}
                      placeholder="e.g. Engineering"
                    />
                    <InputField
                      label="Project(s) Managed"
                      value={data.employee.projectsManaged}
                      onChange={(val) => setData({ ...data, employee: { ...data.employee, projectsManaged: val } })}
                      placeholder="e.g. Project Alpha, Beta"
                    />
                    <div className="grid grid-cols-2 gap-4">
                      <InputField
                        label="Appraisal Type"
                        value={data.employee.appraisalType}
                        onChange={(val) => setData({ ...data, employee: { ...data.employee, appraisalType: val } })}
                        placeholder="e.g. Annual"
                      />
                      <InputField
                        label="Period"
                        value={data.employee.appraisalPeriod}
                        onChange={(val) => setData({ ...data, employee: { ...data.employee, appraisalPeriod: val } })}
                        placeholder="e.g. 2023-2024"
                      />
                    </div>
                  </div>

                  <div className="space-y-6">
                    <h3 className="text-sm font-semibold uppercase tracking-wider text-muted">Reviewer Details</h3>
                    <InputField
                      label="Reviewer Name"
                      value={data.reviewer.name}
                      onChange={(val) => setData({ ...data, reviewer: { ...data.reviewer, name: val } })}
                      placeholder="e.g. Jane Smith"
                    />
                    <InputField
                      label="Reviewer Position"
                      value={data.reviewer.position}
                      onChange={(val) => setData({ ...data, reviewer: { ...data.reviewer, position: val } })}
                      placeholder="e.g. Area Manager"
                    />
                    <InputField
                      label="Department"
                      value={data.reviewer.department}
                      onChange={(val) => setData({ ...data, reviewer: { ...data.reviewer, department: val } })}
                      placeholder="e.g. Operations"
                    />
                    <InputField
                      label="Appraisal Due Date"
                      type="date"
                      value={data.reviewer.appraisalDue}
                      onChange={(val) => setData({ ...data, reviewer: { ...data.reviewer, appraisalDue: val } })}
                    />
                  </div>
                </div>
                <div className="mt-8 flex justify-end">
                  <button
                    type="button"
                    onClick={() => setActiveTab("factors")}
                    className="inline-flex items-center gap-2 rounded-full bg-accent px-5 py-2 text-xs font-bold uppercase tracking-widest text-white shadow-md shadow-accent/30 hover:brightness-110 active:scale-[0.98] transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page"
                  >
                    Done
                    <ChevronRight className="h-4 w-4" />
                  </button>
                </div>
              </Section>
            </motion.div>
          )}

          {activeTab === "factors" && (
            <motion.div
              key="factors"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-6"
            >
              <div className="flex justify-between items-end mb-4 px-2 gap-4">
                <div>
                  <h2 className="text-2xl font-light text-foreground tracking-tight">Performance Appraisal</h2>
                  <p className="text-muted text-sm mt-1 uppercase tracking-widest font-medium">Competency Evaluation</p>
                </div>
                <div className="text-right shrink-0">
                  <span className="text-[10px] font-bold uppercase tracking-[0.2em] text-accent">Section Total</span>
                  <div className="text-4xl font-light text-accent tracking-tighter">{formatScore(stats.performanceScore)}</div>
                </div>
              </div>

              {(() => {
                const idx = factorSlideIndex;
                const factor = data.factors[idx];
                const total = data.factors.length;
                const goPrev = () => {
                  if (idx <= 0) return;
                  setFactorSlideDir(-1);
                  setFactorSlideIndex(idx - 1);
                };
                const goNext = () => {
                  if (idx >= total - 1) return;
                  setFactorSlideDir(1);
                  setFactorSlideIndex(idx + 1);
                };
                const handleDone = () => {
                  if (idx < total - 1) {
                    goNext();
                  } else {
                    setActiveTab("goals");
                  }
                };
                return (
                  <>
                    <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 px-2 mb-4">
                      <p className="text-sm font-semibold text-muted">
                        Factor {factor.id} · {idx + 1} / {total}
                      </p>
                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={goPrev}
                          disabled={idx <= 0}
                          className="inline-flex items-center gap-2 rounded-full border border-border bg-elevated px-4 py-2 text-xs font-bold uppercase tracking-widest text-foreground hover:bg-surface disabled:opacity-40 disabled:pointer-events-none transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page"
                        >
                          <ChevronLeft className="h-4 w-4" />
                          Previous
                        </button>
                        <button
                          type="button"
                          onClick={goNext}
                          disabled={idx >= total - 1}
                          className="inline-flex items-center gap-2 rounded-full border border-border bg-elevated px-4 py-2 text-xs font-bold uppercase tracking-widest text-foreground hover:bg-surface disabled:opacity-40 disabled:pointer-events-none transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page"
                        >
                          Next
                          <ChevronRight className="h-4 w-4" />
                        </button>
                      </div>
                    </div>

                    <AnimatePresence mode="wait" initial={false}>
                      <motion.div
                        key={factor.id}
                        initial={{ opacity: 0, x: 24 * factorSlideDir }}
                        animate={{ opacity: 1, x: 0 }}
                        exit={{ opacity: 0, x: -24 * factorSlideDir }}
                        transition={{ duration: 0.2, ease: "easeOut" }}
                      >
                        <Section
                          className={cn(
                            "relative group transition-all hover:bg-elevated/80",
                            idx % 3 === 0
                              ? "border-l-4 border-accent-warm"
                              : idx % 3 === 1
                                ? "border-l-4 border-accent"
                                : "border-l-4 border-accent-deep",
                          )}
                        >
                          <div className="flex flex-col lg:flex-row gap-8">
                            <div className="flex-1">
                              <div className="flex items-center gap-3 mb-4 flex-wrap">
                                <span className="text-[10px] font-black bg-badge text-badge-foreground px-2 py-1 rounded">
                                  FACTOR {factor.id}
                                </span>
                                <h3 className="text-lg font-bold text-foreground">{factor.title}</h3>
                                <span className="text-xs font-bold text-muted ml-auto tracking-widest">
                                  WGT: {factor.weightage}
                                </span>
                              </div>
                              <p className="text-sm text-muted mb-6 leading-relaxed bg-input p-4 rounded-xl border border-border not-italic">
                                {factor.description}
                              </p>
                              <textarea
                                className="w-full bg-input border border-border rounded-xl p-4 text-sm text-foreground placeholder:text-placeholder focus-visible:ring-2 focus-visible:ring-ring focus-visible:border-accent outline-none transition-all min-h-[100px]"
                                placeholder="Add specific observations or feedback..."
                                value={factor.comments}
                                onChange={(e) => {
                                  const newFactors = [...data.factors];
                                  newFactors[idx].comments = e.target.value;
                                  setData({ ...data, factors: newFactors });
                                }}
                              />
                              <div className="mt-3 flex justify-end">
                                <button
                                  type="button"
                                  onClick={handleDone}
                                  className="inline-flex items-center gap-2 rounded-full bg-accent px-5 py-2 text-xs font-bold uppercase tracking-widest text-white shadow-md shadow-accent/30 hover:brightness-110 active:scale-[0.98] transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page"
                                >
                                  Done
                                  <ChevronRight className="h-4 w-4" />
                                </button>
                              </div>
                            </div>

                            <div className="lg:w-72 space-y-6 flex flex-col justify-center bg-score-panel p-6 rounded-2xl border border-border">
                              <div>
                                <label className="text-[10px] font-bold text-muted uppercase tracking-[0.2em] block mb-3">
                                  Employee Score
                                </label>
                                <ScoreSelect
                                  value={factor.employeeScore}
                                  onChange={(val) => {
                                    const newFactors = [...data.factors];
                                    newFactors[idx].employeeScore = val;
                                    setData({ ...data, factors: newFactors });
                                  }}
                                />
                              </div>
                              <div className="pt-4 border-t border-border">
                                <label className="text-[10px] font-bold text-muted uppercase tracking-[0.2em] block mb-3">
                                  Reviewer Score
                                </label>
                                <ScoreSelect
                                  value={factor.reviewerScore}
                                  onChange={(val) => {
                                    const newFactors = [...data.factors];
                                    newFactors[idx].reviewerScore = val;
                                    setData({ ...data, factors: newFactors });
                                  }}
                                  isReviewer
                                />
                              </div>
                            </div>
                          </div>
                        </Section>
                      </motion.div>
                    </AnimatePresence>
                  </>
                );
              })()}
            </motion.div>
          )}

          {activeTab === "goals" && (
            <motion.div
              key="goals"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-6"
            >
              <div className="flex justify-between items-end mb-4 px-2 gap-4">
                <div>
                  <h2 className="text-2xl font-light text-foreground tracking-tight">Key Achievement Goals</h2>
                  <p className="text-muted text-sm mt-1 uppercase tracking-widest font-medium">Specific objectives evaluation</p>
                </div>
                <div className="text-right shrink-0">
                  <span className="text-[10px] font-bold uppercase tracking-[0.2em] text-accent-warm">Goals Total</span>
                  <div className="text-4xl font-light text-accent-warm tracking-tighter">{formatScore(stats.goalAchievementScore)}</div>
                </div>
              </div>

              {data.goals.map((goal, idx) => (
                <div key={goal.id} id={`goal-${goal.id}`}>
                  <Section className="border-l-4 border-border hover:border-accent transition-all">
                    <div className="space-y-6">
                    <div className="flex items-center gap-3 flex-wrap">
                      <span className="text-[10px] font-black bg-badge text-badge-foreground px-2 py-1 rounded">GOAL {goal.id}</span>
                      <div className="flex-1 min-w-[12rem]">
                        <input
                          className="w-full text-lg font-bold text-foreground bg-transparent border-b border-border focus:border-accent outline-none transition-all placeholder:text-placeholder"
                          placeholder="What was the goal?"
                          value={goal.description}
                          onChange={(e) => {
                            const newGoals = [...data.goals];
                            newGoals[idx].description = e.target.value;
                            setData({ ...data, goals: newGoals });
                          }}
                        />
                      </div>
                      <div className="text-xs font-bold text-muted tracking-widest">WGT: {goal.weightage}</div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-12 gap-6">
                      <div className="md:col-span-8">
                        <textarea
                          className="w-full bg-input border border-border rounded-xl p-4 text-sm text-foreground min-h-[100px] focus-visible:ring-2 focus-visible:ring-ring focus-visible:border-accent outline-none transition-all placeholder:text-placeholder leading-relaxed"
                          placeholder="Results achieved, metrics, or challenges..."
                          value={goal.comments}
                          onChange={(e) => {
                            const newGoals = [...data.goals];
                            newGoals[idx].comments = e.target.value;
                            setData({ ...data, goals: newGoals });
                          }}
                        />
                      </div>
                      <div className="md:col-span-4 flex flex-col justify-center bg-score-panel p-6 rounded-2xl border border-border">
                        <label className="text-[10px] font-bold text-muted uppercase tracking-[0.2em] block mb-3 text-center md:text-left">
                          Reviewer Score
                        </label>
                        <ScoreSelect
                          value={goal.score}
                          onChange={(val) => {
                            const newGoals = [...data.goals];
                            newGoals[idx].score = val;
                            setData({ ...data, goals: newGoals });
                          }}
                          isReviewer
                        />
                      </div>
                    </div>
                    <div className="flex justify-end">
                      <button
                        type="button"
                        onClick={() => {
                          const next = data.goals[idx + 1];
                          if (next) {
                            document.getElementById(`goal-${next.id}`)?.scrollIntoView({ behavior: "smooth", block: "start" });
                            return;
                          }
                          setActiveTab("summary");
                        }}
                        className="inline-flex items-center gap-2 rounded-full bg-accent px-5 py-2 text-xs font-bold uppercase tracking-widest text-white shadow-md shadow-accent/30 hover:brightness-110 active:scale-[0.98] transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page"
                      >
                        Done
                        <ChevronRight className="h-4 w-4" />
                      </button>
                    </div>
                  </div>
                  </Section>
                </div>
              ))}
            </motion.div>
          )}

          {activeTab === "summary" && (
            <motion.div
              key="summary"
              initial={{ opacity: 0, scale: 0.98 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="space-y-8 max-w-5xl mx-auto"
            >
              <div className="bg-gradient-to-br from-accent-warm to-accent rounded-[32px] p-8 sm:p-12 text-center shadow-xl shadow-accent/25 relative overflow-hidden">
                <div className="absolute -top-24 -right-24 w-96 h-96 bg-white/10 rounded-full blur-[100px]" />
                <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-full h-full bg-black/10" />

                <div className="relative z-10 flex flex-col items-center">
                  <h3 className="text-white/80 uppercase text-xs tracking-[0.4em] font-bold mb-8">
                    Computed Growth Matrix Result
                  </h3>
                  <div className="text-7xl sm:text-[120px] font-light leading-none text-white tracking-tighter mb-4">
                    {formatScore(stats.finalScore)}
                  </div>

                  <div className="px-8 py-3 bg-black/25 backdrop-blur-xl rounded-full text-white text-sm font-black uppercase tracking-[0.2em] mb-12 border border-white/25">
                    {stats.ratingCategory}
                  </div>

                  <div className="w-full max-w-md space-y-3">
                    <div className="w-full bg-black/25 h-2 rounded-full overflow-hidden p-0.5 border border-white/15">
                      <motion.div
                        initial={{ width: 0 }}
                        animate={{ width: `${(stats.finalScore / 5) * 100}%` }}
                        className="bg-white h-full rounded-full"
                      />
                    </div>
                    <div className="flex justify-between w-full text-[10px] text-white/70 uppercase font-black tracking-widest">
                      <span>1.0 Needs Imp.</span>
                      <span>5.0 Exceptional</span>
                    </div>
                  </div>
                </div>

                <div className="mt-16 grid grid-cols-2 md:grid-cols-4 gap-4 relative z-10">
                  <div className="bg-black/25 backdrop-blur-md p-4 rounded-2xl border border-white/15">
                    <span className="text-[8px] font-black uppercase text-white/60 tracking-[0.2em]">Perf. (Emp)</span>
                    <div className="text-xl font-bold text-white">{formatScore(stats.employeeFactorScore)}</div>
                  </div>
                  <div className="bg-black/25 backdrop-blur-md p-4 rounded-2xl border border-white/15">
                    <span className="text-[8px] font-black uppercase text-white/60 tracking-[0.2em]">Perf. (Rev)</span>
                    <div className="text-xl font-bold text-white">{formatScore(stats.reviewerFactorScore)}</div>
                  </div>
                  <div className="bg-black/25 backdrop-blur-md p-4 rounded-2xl border border-white/15">
                    <span className="text-[8px] font-black uppercase text-white/60 tracking-[0.2em]">Perf. Score (80%)</span>
                    <div className="text-xl font-bold text-white">{formatScore(stats.performanceScore)}</div>
                  </div>
                  <div className="bg-black/25 backdrop-blur-md p-4 rounded-2xl border border-white/15">
                    <span className="text-[8px] font-black uppercase text-white/60 tracking-[0.2em]">Goals Score (20%)</span>
                    <div className="text-xl font-bold text-white">{formatScore(stats.goalAchievementScore)}</div>
                  </div>
                </div>
              </div>

              <Section title="Final Remarks & Assessment">
                <div className="space-y-8">
                  <CommentField
                    label="Employee Reflection"
                    value={data.comments.employee}
                    onChange={(val) => setData({ ...data, comments: { ...data.comments, employee: val } })}
                  />
                  <CommentField
                    label="Reviewer Summary"
                    value={data.comments.reviewer}
                    onChange={(val) => setData({ ...data, comments: { ...data.comments, reviewer: val } })}
                  />
                  <CommentField
                    label="Leadership Notes (CEO)"
                    value={data.comments.ceo}
                    onChange={(val) => setData({ ...data, comments: { ...data.comments, ceo: val } })}
                  />
                </div>
              </Section>

              <Section title="Signatures">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                  <div className="space-y-4">
                    <InputField
                      label="Employee Name for Signature"
                      value={data.signatures.employee.name}
                      onChange={(val) =>
                        setData({
                          ...data,
                          signatures: {
                            ...data.signatures,
                            employee: { ...data.signatures.employee, name: val },
                          },
                        })
                      }
                    />
                    <InputField
                      type="date"
                      label="Date"
                      value={data.signatures.employee.date}
                      onChange={(val) =>
                        setData({
                          ...data,
                          signatures: {
                            ...data.signatures,
                            employee: { ...data.signatures.employee, date: val },
                          },
                        })
                      }
                    />
                  </div>
                  <div className="space-y-4">
                    <InputField
                      label="Reviewer Name for Signature"
                      value={data.signatures.reviewer.name}
                      onChange={(val) =>
                        setData({
                          ...data,
                          signatures: {
                            ...data.signatures,
                            reviewer: { ...data.signatures.reviewer, name: val },
                          },
                        })
                      }
                    />
                    <InputField
                      type="date"
                      label="Date"
                      value={data.signatures.reviewer.date}
                      onChange={(val) =>
                        setData({
                          ...data,
                          signatures: {
                            ...data.signatures,
                            reviewer: { ...data.signatures.reviewer, date: val },
                          },
                        })
                      }
                    />
                  </div>
                </div>
              </Section>

              <div className="pb-12 pt-4">
                <button
                  type="button"
                  onClick={handleDownload}
                  className="w-full flex items-center justify-center gap-4 rounded-[32px] border-2 border-accent bg-accent p-8 font-black text-lg sm:text-xl text-white shadow-lg shadow-accent/30 hover:brightness-110 active:scale-[0.99] transition-all group overflow-hidden relative focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-page"
                >
                  <Download className="w-8 h-8 text-white group-hover:translate-y-0.5 transition-transform relative z-10" />
                  <span className="relative z-10 tracking-tight">GENERATE & DOWNLOAD OFFICIAL DOCX</span>
                </button>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
        </main>
      </div>

      <div
        className={cn(
          "print-only-sink docx-print-sink",
          previewOpen ? "hidden print:hidden" : "hidden print:block",
        )}
        aria-hidden
      >
        <div ref={printSinkRef} className="docx-preview-host min-h-[1px] w-full bg-white" />
      </div>

      {previewOpen ? (
        <div
          className="print-preview-modal fixed inset-0 z-[100] flex items-end justify-center sm:items-center p-0 sm:p-6"
          role="presentation"
        >
          <button
            type="button"
            className="absolute inset-0 bg-black/60 print:hidden"
            aria-label="Close preview"
            onClick={() => setPreviewOpen(false)}
          />
          <div
            role="dialog"
            aria-modal="true"
            aria-labelledby={previewTitleId}
            className="relative z-10 flex max-h-[92vh] w-full max-w-[56rem] flex-col rounded-t-2xl border border-border bg-surface shadow-2xl sm:rounded-2xl print:max-h-none print:rounded-none print:border-0 print:bg-white print:shadow-none"
          >
            <div className="no-print flex shrink-0 items-center justify-between gap-3 border-b border-border px-4 py-3 sm:px-5">
              <div className="min-w-0">
                <h2 id={previewTitleId} className="text-sm font-bold uppercase tracking-widest text-foreground">
                  Document preview
                </h2>
                <p className="mt-0.5 text-[10px] font-medium uppercase tracking-wider text-muted">
                  Same DOCX as Download — rendered in the browser
                </p>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handlePrint}
                  className="inline-flex items-center gap-2 rounded-full border border-border bg-elevated px-4 py-2 text-xs font-bold uppercase tracking-widest text-foreground hover:bg-surface focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-surface"
                >
                  <Printer className="h-4 w-4" />
                  Print
                </button>
                <button
                  type="button"
                  onClick={() => setPreviewOpen(false)}
                  className="inline-flex items-center justify-center rounded-full border border-border p-2 text-muted hover:text-foreground hover:bg-elevated focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 focus-visible:ring-offset-surface"
                  aria-label="Close"
                >
                  <X className="h-4 w-4" />
                </button>
              </div>
            </div>
            <div className="print-preview-modal-scroll min-h-0 flex-1 overflow-y-auto bg-white px-3 py-4 sm:px-6 sm:py-6">
              {previewDocLoading ? (
                <p className="text-center text-sm text-muted">Generating document preview…</p>
              ) : null}
              {previewDocError ? (
                <p className="text-center text-sm text-red-600" role="alert">
                  {previewDocError}
                </p>
              ) : null}
              <div ref={modalDocxRef} className="docx-preview-host mx-auto w-full max-w-[56rem]" />
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

type SectionProps = Attributes & {
  title?: string;
  children: ReactNode;
  className?: string;
};

function Section({ title, children, className }: SectionProps) {
  return (
    <div
      className={cn(
        "bg-surface rounded-3xl p-8 border border-border shadow-sm backdrop-blur-sm",
        className,
      )}
    >
      {title && (
        <h2 className="text-[10px] font-bold text-muted uppercase tracking-[0.2em] mb-8">{title}</h2>
      )}
      {children}
    </div>
  );
}

function InputField({
  label,
  value,
  onChange,
  placeholder,
  type = "text",
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  placeholder?: string;
  type?: string;
}) {
  const id = useId();
  const isDate = type === "date";
  const inputClass = cn(
    "w-full bg-input border border-border rounded-xl px-4 py-3 text-sm text-foreground placeholder:text-placeholder focus-visible:ring-2 focus-visible:ring-ring focus-visible:border-accent outline-none transition-all",
    isDate && "relative pr-11 date-input-overlap",
  );

  return (
    <div className="space-y-2">
      <label htmlFor={id} className="text-[10px] font-bold text-muted uppercase tracking-widest ml-1">
        {label}
      </label>
      {isDate ? (
        <div className="relative">
          <input
            id={id}
            type="date"
            className={inputClass}
            value={value}
            onChange={(e) => onChange(e.target.value)}
            placeholder={placeholder}
          />
          <Calendar
            className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted"
            aria-hidden
          />
        </div>
      ) : (
        <input
          id={id}
          type={type}
          className={inputClass}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder}
        />
      )}
    </div>
  );
}

function CommentField({ label, value, onChange }: { label: string; value: string; onChange: (v: string) => void }) {
  return (
    <div className="space-y-2">
      <label className="text-[10px] font-bold text-muted uppercase tracking-widest ml-1">{label}</label>
      <textarea
        className="w-full bg-input border border-border rounded-xl px-4 py-4 text-sm text-foreground min-h-[120px] focus-visible:ring-2 focus-visible:ring-ring focus-visible:border-accent outline-none transition-all placeholder:text-placeholder leading-relaxed"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={`Enter ${label.toLowerCase()}...`}
      />
    </div>
  );
}

function ScoreSelect({
  value,
  onChange,
  isReviewer = false,
}: {
  value: number;
  onChange: (v: number) => void;
  isReviewer?: boolean;
}) {
  const scores = [1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5];

  return (
    <div
      className={cn(
        "grid grid-cols-3 gap-1.5 rounded-xl border border-border bg-elevated p-2",
        isReviewer && "bg-score-panel",
      )}
    >
      {scores.map((s) => (
        <button
          type="button"
          key={s}
          onClick={() => onChange(s)}
          className={cn(
            "flex h-10 w-full min-w-0 items-center justify-center rounded-lg text-xs font-black tabular-nums transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-1 focus-visible:ring-offset-elevated",
            value === s
              ? isReviewer
                ? "bg-gradient-to-r from-accent-warm to-accent text-white shadow-md ring-1 ring-white/30"
                : "bg-chip text-chip-text shadow-md ring-1 ring-border"
              : isReviewer
                ? "text-foreground hover:bg-surface"
                : "text-muted hover:bg-surface hover:text-foreground",
          )}
        >
          {s}
        </button>
      ))}
    </div>
  );
}
