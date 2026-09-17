import { useEffect, useState } from 'react'
import { GanttScreen } from './gantt/GanttScreen'
import { GanttToolbar } from './gantt/GanttToolbar'
import { SettingsFlyout } from './components/SettingsFlyout'
import { ImportModal } from './components/ImportModal'
import { ExportMenu, type DeckFormat, type ExportFormat } from './components/ExportMenu'
import { Settings } from 'lucide-react'
import { PlanNotice } from './components/PlanNotice'
import { exportTimelineToPptx } from './export/pptxExporter'
import { exportTimelineToPdf } from './export/pdfExporter'
import { downloadPlanCsv } from './export/planCsv'
import { exportPlanToJsonFile } from './import/planJson'
import { buildExportFilename } from './export/dateScale'
import { flushedActivePlan, useTimelineStore } from './store/timelineStore'
import { buttonBaseClass } from './components/systemUi';

function App() {
  const loadPlans = useTimelineStore((state) => state.loadPlans)
  const items = useTimelineStore((state) => state.items)
  const exportOptions = useTimelineStore((state) => state.exportOptions)
  const comments = useTimelineStore((state) => state.comments)

  // What the deck is made of: the whole plan. Nothing on this screen narrows
  // it — a branch worth its own deck has its own plan now (see
  // createPlanFromBranch). What a deck leaves out is said per task, with
  // "Exclude from export", and the exporters read that themselves.
  const [isSettingsOpen, setIsSettingsOpen] = useState(false)
  // Only for the Export menu's JSON row, which is disabled until there is a
  // plan record to write. Everything else about that file is read at the click
  // (see handleExport), not subscribed to.
  const activePlanId = useTimelineStore((state) => state.activePlanId)
  const [isImportOpen, setIsImportOpen] = useState(false)

  useEffect(() => {
    void loadPlans()
  }, [loadPlans])

  const runExport = (format: DeckFormat) => {
    const fileName = buildExportFilename(exportOptions.exportTimeframe, format)
    const exportTimeline = format === 'pptx' ? exportTimelineToPptx : exportTimelineToPdf
    // Always 'full': every exportable task reaches the deck, paged across as
    // many overview slides as that takes. A plan that fits on one slide still
    // gets one — 'full' and 'compact' only differ once there are more tasks
    // than fit, and at that point silently dropping the rest is not a thing
    // an export should do on its own.
    void exportTimeline(items, exportOptions, comments, fileName, 'full')
  }

  const handleExport = (format: ExportFormat) => {
    // A table has no slides to page across at all: every exportable task is a
    // row, however many there are. The filename is built by the same rule as
    // the other two.
    if (format === 'csv') {
      downloadPlanCsv(items, buildExportFilename(exportOptions.exportTimeframe, 'csv'))
      return
    }

    // The plan itself rather than a rendering of it, so it is the plan record
    // that is written — with the working copy flushed into it first, or the
    // file would be missing every edit made since this plan was opened while
    // the three formats above all read what is on screen.
    if (format === 'json') {
      const plan = flushedActivePlan(useTimelineStore.getState())
      if (plan) exportPlanToJsonFile(plan)
      return
    }

    runExport(format)
  }

  // The shell is `100vh` tall, and `100dvh` below the mobile breakpoint. On
  // iOS Safari the layout viewport keeps running under the address bar, so a
  // screen measured in `vh` puts its last row — here the chart's own bottom
  // edge, and the create lane along it — behind that bar until it retracts.
  // The two units are the same number on a desktop, and the utility is scoped
  // to `max-md` anyway: nothing above 768px changes.
  return (
    <div className="flex h-screen flex-col overflow-hidden bg-base-background max-md:h-[100dvh]">
      {/* One header for the whole app. The plan's own controls are the
          handoff's; this app's four actions and its view switch join them at
          the right of the top row, because there is nowhere else for them to
          live once the screen is the plan. */}
      <GanttToolbar
        actions={
          <>
            {/* The two things the app does to a whole plan, both said as the
                verb. Which file an export produces is a question for the menu
                behind it, not for the header. */}
            <button
              type="button"
              onClick={() => setIsImportOpen(true)}
              // Taller and tighter below the breakpoint: 40px is a target a
              // thumb can find, and the two px it gives back go to the plan's
              // name at the other end of the row.
              className={buttonBaseClass(
                'outline',
                'h-8 whitespace-nowrap px-3 text-xs font-semibold max-md:h-10 max-md:px-2.5',
              )}
            >
              Import
            </button>
            <ExportMenu onExport={handleExport} hasSavedPlan={activePlanId !== null} />
            {/* The export settings, beside the button whose files they
                govern — what to include, in what order, over what window.
                No label: a gear is the one glyph that needs none, and the
                row's width is the scarcest thing in this header. */}
            <button
              type="button"
              onClick={() => setIsSettingsOpen(true)}
              title="Export settings"
              aria-label="Export settings"
              className={buttonBaseClass('outline', 'h-8 w-8 flex-none px-0 max-md:h-10 max-md:w-10')}
            >
              <Settings size={15} strokeWidth={2} aria-hidden="true" />
            </button>
          </>
        }
      />

      <main className="flex min-h-0 flex-1 flex-col">
        {/* Sits with the plan it describes, above the chart whose order the
            repair changed — and renders nothing at all when there was none. */}
        <div className="flex-none empty:hidden [&>*]:mx-6 [&>*]:mt-4 max-md:[&>*]:mx-3">
          <PlanNotice />
        </div>

        {/* The one screen this app has. It runs edge to edge and owns its own
            scrolling: its canvas is one scroll container that has to reach the
            window's edges to be worth scrolling. */}
        <GanttScreen />
      </main>

      {isSettingsOpen && <SettingsFlyout onClose={() => setIsSettingsOpen(false)} />}

      {isImportOpen && <ImportModal onClose={() => setIsImportOpen(false)} />}

    </div>
  )
}

export default App
