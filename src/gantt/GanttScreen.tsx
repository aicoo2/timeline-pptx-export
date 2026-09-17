import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ChevronRight, Eye, EyeOff, Pencil, Plus, Trash2, X } from 'lucide-react';
import { ContextMenu, type ContextMenuAction } from '../components/ContextMenu';
import { buttonBaseClass } from '../components/systemUi';
import { useTimelineStore } from '../store/timelineStore';
import { useIsMobile } from '../utils/useIsMobile';
import { buildNewTask } from '../utils/newTask';
import { buildBranchColors } from '../utils/branchColors';
import { buildDepthMap } from '../utils/barNesting';
import { buildBarStyles } from './barColor';
import { useSiteTheme } from '../utils/siteTheme';
import { progressForStatus } from '../utils/progressForStatus';
import { useScrollPanes } from './useScrollPanes';
import type { TaskStatus } from '../types/timeline';
import {
  clampListWidth,
  HEADER_HEIGHT_PX,
  LIST_RESIZE_HANDLE_PX,
  MAX_LIST_WIDTH_PX,
  MIN_BODY_HEIGHT_PX,
  MIN_LIST_WIDTH_PX,
  ADD_ROW_HEIGHT_PX,
  ROW_HEIGHT_PX,
  TODAY_SCROLL_LEAD_PX,
  type Span,
} from './geometry';
import {
  buildHeaderCells,
  COLUMN_WIDTH_PX,
  formatDayLabel,
  isoAtIndex,
  planRange,
  weekendStarts,
} from './scale';
import { confirmTaskDeletion, deleteTaskBranch } from './deleteTask';
import { childrenOf, isGroup, isSubtask } from './rollup';
import { visibleRows } from './rows';
import { previewSpans, type DragState } from './drag';
import { STATUS_CYCLE, STATUS_LABEL } from './tone';
import { TimelineHeader } from './TimelineHeader';
import { TaskList } from './TaskList';
import { TimelineBody } from './TimelineBody';
import { EditTaskPanel } from './EditTaskPanel';
import { useGanttViewStore } from './viewStore';

/** How long after a pointerup a drag keeps suppressing the click it would
 * otherwise fire. Long enough for the click event to have been and gone,
 * short enough that the next real click lands. */
const CLICK_SUPPRESS_MS = 30;

/** What a sub-task is called for the moment between being created and being
 * named. The name field opens on it immediately, so this is what stands if
 * the rename is abandoned rather than a placeholder anyone types over. */
const NEW_SUBTASK_LABEL = 'New sub-task';

/** How far a finger has to travel on the drawer before the gesture is read as
 * a swipe rather than as a scroll or a tap — and, once it is a swipe, how far
 * left it has to go for the release to close the drawer rather than snap it
 * back. The first is the same 3px the canvas's own pan uses to tell a press
 * from a drag; the second is a third of the narrowest drawer, which is far
 * enough that no glancing sideways movement during a scroll reaches it. */
const SWIPE_AXIS_PX = 3;
const SWIPE_CLOSE_PX = 80;

/** Where the opening date sits in the zone below the breakpoint, as a share
 * of its width — the same three-tenths TODAY_SCROLL_LEAD_PX is on a desktop
 * timeline, said as a proportion because at 375px a fixed 300px is not a lead
 * but the whole screen. */
const MOBILE_TODAY_LEAD_SHARE = 0.3;

/** The plan screen: toolbar, then one scroll container holding the task list,
 * the period header and the timeline, then the Edit Task panel.
 *
 * The scroll container is the whole of the alignment strategy. It holds a CSS
 * grid — `320px <timeline>` by `56px <body>` — with the header stuck to its
 * top, the list stuck to its left and the corner block stuck to both. A row's
 * line in the list and its bar in the timeline are the same grid row, so they
 * cannot drift apart; nothing here mirrors one scroll onto another. */
export function GanttScreen() {
  const items = useTimelineStore((state) => state.items);
  const updateItem = useTimelineStore((state) => state.updateItem);
  const addItem = useTimelineStore((state) => state.addItem);
  const toggleIncludeInExportCascade = useTimelineStore((state) => state.toggleIncludeInExportCascade);
  const activePlanId = useTimelineStore((state) => state.activePlanId);

  const scale = useGanttViewStore((state) => state.scale);
  const listWidth = useGanttViewStore((state) => state.listWidth);
  const setListWidth = useGanttViewStore((state) => state.setListWidth);
  const collapsed = useGanttViewStore((state) => state.collapsed);
  const selectedId = useGanttViewStore((state) => state.selectedId);
  const select = useGanttViewStore((state) => state.select);
  const toggleCollapsed = useGanttViewStore((state) => state.toggleCollapsed);
  const collapseBranch = useGanttViewStore((state) => state.collapseBranch);
  const beginRename = useGanttViewStore((state) => state.beginRename);

  // Below the breakpoint the task list is a drawer over the chart rather than
  // a column beside it. Every mobile-only branch on this screen hangs off this
  // one flag, and above the breakpoint none of them is taken.
  const isMobile = useIsMobile();
  const isDrawerOpen = useGanttViewStore((state) => state.isTaskDrawerOpen);
  const setDrawerOpen = useGanttViewStore((state) => state.setTaskDrawerOpen);
  const drawerRef = useRef<HTMLDivElement | null>(null);

  // The row a right-click opened a menu on, and where the pointer was.
  const [menu, setMenu] = useState<{ id: string; x: number; y: number } | null>(null);
  const [drag, setDrag] = useState<DragState | null>(null);
  // Set the moment a drag snaps a whole column, and cleared shortly after the
  // pointer comes up — the flag that tells "pressed a bar" from "moved a bar",
  // so releasing after a drag doesn't also open the panel. The drawer's
  // swipe-to-close raises it for the same reason: a gesture that moved is not
  // a tap on whatever it happened to start over.
  const movedRef = useRef(false);

  // One "now" per mount. Re-reading the clock on every render would let the
  // today band and the canvas origin change under a drag.
  const today = useMemo(() => new Date(), []);
  const { minDate, totalDays, todayIndex, firstTaskIndex, lastTaskIndex } = useMemo(
    () => planRange(items, today),
    [items, today],
  );
  const columnWidth = COLUMN_WIDTH_PX[scale];

  const rows = useMemo(() => visibleRows(items, { collapsed }), [items, collapsed]);

  const bodyHeight = Math.max(rows.length * ROW_HEIGHT_PX + ADD_ROW_HEIGHT_PX, MIN_BODY_HEIGHT_PX);

  // The three panes and the one scroller between them. The column width is
  // all the hook needs from here: it is how it recognises a scale change.
  const panes = useScrollPanes({ columnWidth });

  // The width a drag on the seam is currently proposing, or null when no drag
  // is under way. Held here rather than written straight to the store because
  // the store is persisted: committing on pointerup is one localStorage write
  // per drag instead of one per pixel moved.
  const [dragWidth, setDragWidth] = useState<number | null>(null);
  const listResizeRef = useRef<{ startX: number; startWidth: number } | null>(null);
  const shownListWidth = dragWidth ?? listWidth;
  const isResizingList = dragWidth !== null;

  const beginListResize = (event: React.PointerEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    listResizeRef.current = { startX: event.clientX, startWidth: listWidth };
    setDragWidth(listWidth);
    event.currentTarget.setPointerCapture(event.pointerId);
  };

  const moveListResize = (event: React.PointerEvent<HTMLDivElement>) => {
    const start = listResizeRef.current;
    if (!start) return;
    setDragWidth(clampListWidth(start.startWidth + (event.clientX - start.startX)));
  };

  const endListResize = (event: React.PointerEvent<HTMLDivElement>) => {
    if (!listResizeRef.current) return;
    listResizeRef.current = null;
    event.currentTarget.releasePointerCapture(event.pointerId);
    setDragWidth((width) => {
      if (width !== null) setListWidth(width);
      return null;
    });
  };

  // The seam is draggable without a pointer too — 16px a press, which is a
  // visible step without being a jump.
  const nudgeListResize = (event: React.KeyboardEvent<HTMLDivElement>) => {
    const step = event.key === 'ArrowLeft' ? -16 : event.key === 'ArrowRight' ? 16 : 0;
    if (step === 0) return;
    event.preventDefault();
    setListWidth(listWidth + step);
  };

  // The drag holds the col-resize cursor across the whole window, not just the
  // 9px strip the pointer started on.
  useEffect(() => {
    if (!isResizingList) return;
    document.body.classList.add('gantt-col-resizing');
    return () => document.body.classList.remove('gantt-col-resizing');
  }, [isResizingList]);

  // A drawer only exists below the breakpoint, so leaving it puts the flag
  // back — otherwise a window dragged wide and narrow again comes back to a
  // drawer standing open over the chart that nobody opened.
  useEffect(() => {
    if (!isMobile) setDrawerOpen(false);
  }, [isMobile, setDrawerOpen]);

  /** A phone shows one layer at a time, so opening the panel puts the drawer
   * away.
   *
   * Below the breakpoint both of them are the screen — the drawer is 82vw of
   * it and the panel is all of it — and the panel is on top. Left standing,
   * the drawer was not a second view but a trap: adding a task from its add
   * row selects that task, so the panel opened over the very field that was
   * still holding focus, and the next thing typed went into an input nobody
   * could see. Escape went to that hidden field first, so it took two presses
   * to get out and the second took the drawer with it, and there is nothing to
   * press outside a panel that covers the screen — which left the ✕ as the
   * only exit anyone could find.
   *
   * So the drawer closes as the panel opens, and it does not come back when
   * the panel does: opening the list is one tap on Tasks, and reopening it
   * over a plan someone just came back to look at would be the same
   * presumption in the other direction. The round trip is open the list, add,
   * fill in, close — then open the list again for the next one.
   *
   * Above the breakpoint there is no drawer and this does nothing: the list is
   * a column and the panel sits beside it, both visible at once. */
  useEffect(() => {
    if (isMobile && selectedId !== null) setDrawerOpen(false);
  }, [isMobile, selectedId, setDrawerOpen]);

  // Escape closes it, as it closes every other layer in this app. A phone
  // rarely has the key; a phone-sized window on a desktop always does.
  //
  // Except over a field, where Escape already means something: the rename and
  // the add row both discard their draft on it, and one key cannot both cancel
  // an edit and take away the column it was being made in.
  useEffect(() => {
    if (!isMobile || !isDrawerOpen) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      if ((event.target as Element | null)?.closest('input, textarea, select')) return;
      setDrawerOpen(false);
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [isMobile, isDrawerOpen, setDrawerOpen]);

  /** The swipe that puts the drawer away.
   *
   * One press, two possible gestures, and which one it is decided by the first
   * few pixels: mostly sideways is a swipe, mostly up or down is a scroll. The
   * scroll is not handled here at all — the drawer carries `touch-action:
   * pan-y`, so the browser keeps the vertical axis and scrolls the list itself
   * (see useScrollPanes for where that offset goes), and this only ever claims
   * the horizontal one.
   *
   * The pointer is followed on `window` rather than captured, which is what
   * lets a press that turned out to be a tap still reach the row underneath:
   * capturing would retarget the click to the drawer and every tap on a name
   * would land nowhere. A swipe that did move is kept off the row by the same
   * `movedRef` the bar drags use. */
  const beginDrawerSwipe = (event: React.PointerEvent<HTMLDivElement>) => {
    if (!isMobile || !isDrawerOpen) return;
    const drawer = drawerRef.current;
    if (!drawer) return;

    const startX = event.clientX;
    const startY = event.clientY;
    let axis: 'x' | 'y' | null = null;
    let dx = 0;

    const onMove = (moveEvent: PointerEvent) => {
      dx = moveEvent.clientX - startX;
      const dy = moveEvent.clientY - startY;
      if (axis === null) {
        if (Math.abs(dx) < SWIPE_AXIS_PX && Math.abs(dy) < SWIPE_AXIS_PX) return;
        axis = Math.abs(dx) > Math.abs(dy) ? 'x' : 'y';
        if (axis === 'x') {
          drawer.dataset.swiping = 'true';
          movedRef.current = true;
        }
      }
      if (axis !== 'x') return;
      // Leftward only. The drawer came from the left edge and that is the only
      // way back; dragging it right would open a gap the chart cannot fill.
      drawer.style.transform = `translateX(${Math.min(0, dx)}px)`;
    };

    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      window.removeEventListener('pointercancel', onUp);
      delete drawer.dataset.swiping;
      // Handed back to the class, which is where both resting positions and
      // the transition between them live.
      drawer.style.transform = '';
      if (axis === 'x') {
        if (dx <= -SWIPE_CLOSE_PX) setDrawerOpen(false);
        window.setTimeout(() => {
          movedRef.current = false;
        }, CLICK_SUPPRESS_MS);
      }
    };

    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
    window.addEventListener('pointercancel', onUp);
  };

  // How many day columns are actually drawn. At least the plan's own span,
  // and at least enough to fill the timeline zone — otherwise a plan shorter
  // than the zone leaves a bare strip beside a grid that stops in mid-air.
  // The prototype never meets this: its canvas is a fixed 133 days, wider
  // than any window it was drawn at. A canvas derived from the plan has to
  // say so itself.
  const renderedDays = Math.max(totalDays, Math.ceil(panes.viewportWidth / columnWidth));
  const canvasWidth = renderedDays * columnWidth;

  // Clamped against what is drawn rather than against the plan's own extent:
  // a bar can be dragged anywhere on the canvas the eye can see.
  const spans = useMemo(
    () => previewSpans(items, minDate, renderedDays, drag),
    [items, minDate, renderedDays, drag],
  );
  const headerCells = useMemo(
    () => buildHeaderCells(minDate, renderedDays, scale),
    [minDate, renderedDays, scale],
  );
  const weekends = useMemo(
    () => weekendStarts(minDate, renderedDays, scale),
    [minDate, renderedDays, scale],
  );

  // Colour is branch, not status — the rule the exported deck draws by, shared
  // with it through buildBranchColors so the two can't drift. Both maps are
  // built from the whole plan, so excluding a task from the export doesn't
  // shift the palette under its neighbours.
  // Colour is branch, not status (above) — but a *tint* is that colour mixed
  // with the canvas, so the canvas is an input and the theme decides it. This
  // is the one thing on the plan screen a token swap cannot carry.
  const theme = useSiteTheme();
  const barStyleById = useMemo(
    () => buildBarStyles(buildBranchColors(items), buildDepthMap(items), theme),
    [items, theme],
  );

  const dateRangeById = useMemo(
    () =>
      new Map(
        rows.map((row) => {
          const span = spans.get(row.item.id);
          if (!span) return [row.item.id, ''];
          return [
            row.item.id,
            `${formatDayLabel(minDate, span.start)} – ${formatDayLabel(minDate, span.start + span.len - 1)}`,
          ];
        }),
      ),
    [rows, spans, minDate],
  );

  const statusLabelById = useMemo(
    () => new Map(rows.map((row) => [row.item.id, STATUS_LABEL[row.status]])),
    [rows],
  );

  // How far into the zone the opening date lands. The constant is measured
  // against a desktop timeline — 300px of a ~950px zone — and on a 375px
  // screen the same 300px would put today four fifths of the way across,
  // against the right edge, with the whole of the plan's future off screen.
  // Below the breakpoint it is read as the share of the zone it always was.
  const todayLead = isMobile
    ? Math.round(panes.viewportWidth * MOBILE_TODAY_LEAD_SHARE)
    : TODAY_SCROLL_LEAD_PX;

  const scrollToToday = useCallback(
    (behavior: ScrollBehavior) => panes.scrollToDay(todayIndex, columnWidth, todayLead, behavior),
    [panes, todayIndex, columnWidth, todayLead],
  );

  // The toolbar's Today button, which lives outside this component now (see
  // requestToday). Skips the very first run so it does not fight the mount
  // scroll below with a smooth animation over the top of it.
  const todayNonce = useGanttViewStore((state) => state.todayNonce);
  const seenNonce = useRef(todayNonce);
  useEffect(() => {
    if (seenNonce.current === todayNonce) return;
    seenNonce.current = todayNonce;
    scrollToToday('smooth');
  }, [todayNonce, scrollToToday]);

  // Where the view opens.
  //
  // Today, when today is somewhere inside the plan — what is happening now is
  // what the view is usually opened to find. But a plan is not always around
  // today: a file can carry a root left behind in 2025 and another one out in
  // 2027, and then the canvas is two years wide with almost all of it empty.
  // Opening such a plan on today would land on bare grid, so it opens on its
  // first task instead — the first thing there is to see.
  const opensOn = todayIndex >= firstTaskIndex && todayIndex <= lastTaskIndex ? todayIndex : firstTaskIndex;
  // Once per plan, not once per mount: switching plans, or importing one over
  // the top of another, lands on a canvas of a different width where the old
  // scroll offset means a different date — usually a stretch of empty grid.
  const openedPlanRef = useRef<string | null>(null);
  useEffect(() => {
    // Waits for the zone to have been measured, since the target is worked
    // out against its width.
    if (panes.viewportWidth === 0) return;
    const planKey = activePlanId ?? '';
    if (openedPlanRef.current === planKey) return;
    openedPlanRef.current = planKey;
    panes.scrollToDay(opensOn, columnWidth, todayLead, 'auto');
  }, [panes, opensOn, columnWidth, activePlanId, todayLead]);

  /** Writes a finished drag onto the plan.
   *
   * One path for every bar. A phase has its own two dates like any other task,
   * so it is moved and resized like any other task — and it moves alone: its
   * sub-tasks keep the dates they had, which is what makes the two editable
   * independently. */
  const commitDrag = useCallback(
    (finished: DragState) => {
      const span = spans.get(finished.id);
      if (!span) return;
      const startIso = isoAtIndex(minDate, span.start);
      const endIso = isoAtIndex(minDate, span.start + span.len - 1);

      if (finished.mode === 'move') updateItem(finished.id, { start: startIso, end: endIso });
      else if (finished.mode === 'start') updateItem(finished.id, { start: startIso });
      else updateItem(finished.id, { end: endIso });
    },
    [minDate, spans, updateItem],
  );

  const commitRef = useRef(commitDrag);
  commitRef.current = commitDrag;

  /** Starts a move or a resize. The pointer is followed on `window`, not on
   * the bar, so a fast drag that outruns the cursor keeps its grip.
   *
   * It ends three ways, and only the first writes anything. A release commits
   * the days the bar was carried. A `pointercancel` puts the bar back where it
   * was — the browser sends one when it takes the pointer over for a scroll,
   * which is what a finger dragged across a bar on a phone actually is, and
   * without this the drag never ended: the listeners stayed on `window`, the
   * preview stayed on the canvas, and the bar went on being drawn at dates the
   * task does not hold. And a drag of zero days commits nothing either way. */
  const beginDrag = useCallback(
    (id: string, event: React.PointerEvent, mode: DragState['mode']) => {
      event.preventDefault();
      event.stopPropagation();

      const startX = event.clientX;
      movedRef.current = false;
      let latest: DragState = { id, mode, days: 0 };

      const detach = () => {
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onUp);
        window.removeEventListener('pointercancel', onCancel);
      };

      const onMove = (moveEvent: PointerEvent) => {
        const days = Math.round((moveEvent.clientX - startX) / columnWidth);
        if (days !== 0) movedRef.current = true;
        latest = { id, mode, days };
        setDrag(latest);
      };

      const onUp = () => {
        detach();
        if (latest.days !== 0) commitRef.current(latest);
        setDrag(null);
        window.setTimeout(() => {
          movedRef.current = false;
        }, CLICK_SUPPRESS_MS);
      };

      const onCancel = () => {
        detach();
        setDrag(null);
        movedRef.current = false;
      };

      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', onUp);
      window.addEventListener('pointercancel', onCancel);
    },
    [columnWidth],
  );

  const cycleStatus = (id: string) => {
    const item = items.find((candidate) => candidate.id === id);
    if (!item) return;
    // A group's status is its children's; there is nothing here to cycle.
    if (isGroup(items, id)) return;

    const current: TaskStatus = item.status ?? 'todo';
    const next = STATUS_CYCLE[(STATUS_CYCLE.indexOf(current) + 1) % STATUS_CYCLE.length];
    // Done means finished and not started means untouched, so both move the
    // percentage with them; the other two say nothing about it and leave it
    // alone. One rule, shared with the panel's status field.
    updateItem(id, { status: next, progress: progressForStatus(next) ?? item.progress });
  };

  /** A task drawn on the timeline's create lane, over the days the drag
   * covered.
   *
   * A root task, "not started", like anything the add row makes — the lane
   * runs under the whole plan rather than inside any one branch's area, so
   * there is no parent for it to belong to, and it is appended after the last
   * row exactly as the add row's tasks are. Everything else is
   * `buildNewTask`'s, so a task drawn here cannot quietly differ from a task
   * typed in the column.
   *
   * The new row is deliberately *not* selected: the Edit Task panel is 348px
   * that would open across the timeline and re-measure the canvas between one
   * drawn task and the next, and drawing tasks is usually drawing several. The
   * bar appearing where the ghost was is confirmation enough. */
  const createTaskOnLane = (span: Span, name: string) => {
    addItem(
      buildNewTask({
        label: name,
        start: isoAtIndex(minDate, span.start),
        end: isoAtIndex(minDate, span.start + span.len - 1),
        status: 'todo',
      }),
    );
  };

  const addTask = (name: string) => {
    const task = buildNewTask({
      label: name,
      start: isoAtIndex(minDate, todayIndex),
      end: isoAtIndex(minDate, todayIndex + 4),
      status: 'todo',
    });
    addItem(task);
    select(task.id);
  };

  /** A sub-task under `parentId`, spanning its parent.
   *
   * The defaults are the ones the bar menu used before this screen was
   * rebuilt: the parent's own span (which is always a valid range), status
   * "not started", nothing else set. What has changed is where the naming
   * happens — the old menu collected a name, two dates and a status in a
   * popup before creating anything, and this screen has an inline name field
   * on every row already, so the task is created first and named in place.
   * The span comes from the drawn spans rather than the parent's stored pair
   * so that a sub-task created mid-drag starts where the bar is being dropped;
   * with the date roll-up gone the two are otherwise the same dates. The
   * sub-task is free to leave that span the moment it is dragged — nothing
   * pulls the parent after it. */
  const addSubtask = (parentId: string) => {
    const parent = items.find((item) => item.id === parentId);
    if (!parent) return null;
    // One level deep and no further, because the client asked for exactly two
    // — a task and its sub-tasks — from this screen. The controls that reach
    // here (the row's "+" and the menu's row) are already absent on a
    // sub-task; the guard is here so the rule holds at the one place a
    // sub-task is actually made, rather than in each control that offers to
    // make one. An imported plan may be as deep as its file is: the cap is on
    // creating, not on holding. See rollup.ts's isSubtask and
    // docs/nesting-depth.md.
    if (isSubtask(items, parentId)) return null;

    const span = spans.get(parentId);
    const task = buildNewTask(
      {
        label: NEW_SUBTASK_LABEL,
        start: span ? isoAtIndex(minDate, span.start) : parent.start,
        end: span ? isoAtIndex(minDate, span.start + span.len - 1) : parent.end,
        status: 'todo',
      },
      { parentId },
    );
    addItem(task);
    // A collapsed parent would swallow the row that is about to be renamed.
    if (collapsed[parentId]) toggleCollapsed(parentId);
    beginRename(task.id, task.label);
  };

  /** This item and every group under it — what "the branch" means to the
   * action that folds one away. */
  const branchGroupIds = (id: string): string[] => {
    const children = childrenOf(items, id);
    if (children.length === 0) return [];
    return [id, ...children.flatMap((child) => branchGroupIds(child.id))];
  };

  const deleteTask = (id: string) => {
    const item = items.find((candidate) => candidate.id === id);
    if (!item) return;
    if (!confirmTaskDeletion(items, item)) return;
    deleteTaskBranch(items, id);
    // Only when it was this row that was open: deleting from the menu while
    // the panel shows a different task leaves that panel alone.
    if (selectedId === id) select(null);
  };

  /** The rows a right-click offers, on the name and on the bar alike. Five at
   * most, and two of them only where they mean anything: a task with no
   * sub-tasks has no branch to fold away, and a task that is itself a
   * sub-task takes no sub-tasks of its own.
   *
   * Making a plan out of a branch is not among them, and is not offered
   * anywhere else either: plans are made from "New plan" and nothing else. */
  const menuActions = (id: string): ContextMenuAction[] => {
    const isBranch = isGroup(items, id);
    const isIncluded = items.find((candidate) => candidate.id === id)?.includeInExport !== false;
    return [
      ...(isSubtask(items, id)
        ? []
        : [
            {
              label: 'Add sub-task',
              icon: <Plus size={14} strokeWidth={2} aria-hidden="true" />,
              onSelect: () => addSubtask(id),
            },
          ]),
      {
        label: 'Rename',
        icon: <Pencil size={14} strokeWidth={2} aria-hidden="true" />,
        onSelect: () => {
          const item = items.find((candidate) => candidate.id === id);
          if (item) beginRename(item.id, item.label);
        },
      },
      ...(isBranch
        ? [
            {
              label: 'Hide sub-tasks',
              icon: <ChevronRight size={14} strokeWidth={2} aria-hidden="true" />,
              onSelect: () => collapseBranch(branchGroupIds(id)),
            },
          ]
        : []),
      {
        // Named for what it will do, not for the state it is in. The whole
        // branch travels together, because a phase in the deck without its
        // tasks — or tasks without the phase they sit under — is not a thing
        // anyone means to export.
        label: isIncluded
          ? `Exclude ${isBranch ? 'branch ' : ''}from export`
          : `Include ${isBranch ? 'branch ' : ''}in export`,
        icon: isIncluded ? (
          <EyeOff size={14} strokeWidth={2} aria-hidden="true" />
        ) : (
          <Eye size={14} strokeWidth={2} aria-hidden="true" />
        ),
        onSelect: () => toggleIncludeInExportCascade(id),
      },
      {
        label: 'Delete',
        icon: <Trash2 size={14} strokeWidth={2} aria-hidden="true" />,
        destructive: true,
        onSelect: () => deleteTask(id),
      },
    ];
  };

  const openMenu = (id: string, event: React.MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    setMenu({ id, x: event.clientX, y: event.clientY });
  };

  const dragSpan = drag ? spans.get(drag.id) : undefined;
  const dragLabel = dragSpan
    ? `${formatDayLabel(minDate, dragSpan.start)} → ${formatDayLabel(minDate, dragSpan.start + dragSpan.len - 1)}  (${dragSpan.len}d)`
    : '';

  const selectedItem = items.find((item) => item.id === selectedId) ?? null;
  // Resolved against the plan rather than trusted: the row a menu is open on
  // can be deleted by that very menu.
  const menuItem = menu === null ? null : (items.find((item) => item.id === menu.id) ?? null);

  return (
    <div style={{ display: 'flex', flex: 1, minHeight: 0, background: 'var(--gantt-surface)' }}>
      {/* The frame. Two tracks by two: the header's height is a constant, the
          list's width is the one the seam was last dragged to, and the
          timeline zone takes whatever is left. Nothing here scrolls — the
          frame is what the three panes inside it scroll *within*, which is
          what keeps the zone the same width at every scale and keeps the
          scrollbar in view.

          Widening the list narrows the body pane, and the body is what
          `useScrollPanes` measures — so the canvas's own width, and with it
          the timeline's right edge, follows the drag without being told. */}
      <div
        style={{
          flex: 1,
          minWidth: 0,
          minHeight: 0,
          position: 'relative',
          display: 'grid',
          // One column below the breakpoint: the chart has the whole width and
          // the task list, which is a drawer there, is taken out of the grid
          // and laid over the top of it.
          gridTemplateColumns: isMobile ? 'minmax(0, 1fr)' : `${shownListWidth}px minmax(0, 1fr)`,
          gridTemplateRows: `${HEADER_HEIGHT_PX}px minmax(0, 1fr)`,
          overflow: 'hidden',
        }}
      >
        {/* The seam, as a grab strip. It sits over the rule the list pane
            draws rather than replacing it: a 1px border is the right line to
            look at and the wrong one to aim at, so the target is 9px wide
            and centred on it, and stays invisible until hovered.

            A sibling of the panes rather than a child of either, because it
            spans both grid rows — the corner above and the list below — and
            because being on top is what keeps its press away from the body's
            grab-to-pan. */}
        {/* Only where there is a seam to drag. Below the breakpoint the two
            zones do not share an edge — one is over the other — and the
            drawer's width is not a thing anyone sets. */}
        {!isMobile && (
        <div
          role="separator"
          aria-orientation="vertical"
          aria-label="Resize the task column"
          aria-valuenow={shownListWidth}
          aria-valuemin={MIN_LIST_WIDTH_PX}
          aria-valuemax={MAX_LIST_WIDTH_PX}
          tabIndex={0}
          className="gantt-col-resize"
          data-dragging={isResizingList}
          onPointerDown={beginListResize}
          onPointerMove={moveListResize}
          onPointerUp={endListResize}
          onPointerCancel={endListResize}
          onKeyDown={nudgeListResize}
          style={{
            position: 'absolute',
            top: 0,
            bottom: 0,
            // Centred on the 1px border, not on the column's edge: the
            // border occupies the pixel before `shownListWidth`, so the
            // strip starts half its width plus that pixel to the left.
            left: shownListWidth - Math.ceil(LIST_RESIZE_HANDLE_PX / 2),
            width: LIST_RESIZE_HANDLE_PX,
            zIndex: 20,
            cursor: 'col-resize',
            touchAction: 'none',
          }}
        />
        )}

        {/* The corner. It follows neither axis, so it is the one pane that is
            simply a box. There is no corner below the breakpoint — with one
            column the ruler starts at the left edge — and the 56px it occupied
            becomes the drawer's own head instead, which is what keeps the
            first name level with the first bar. */}
        {!isMobile && (
        <div
          style={{
            gridColumn: 1,
            gridRow: 1,
            background: 'var(--gantt-surface)',
            borderRight: '1px solid var(--gantt-rule-strong)',
            borderBottom: '1px solid var(--gantt-rule-strong)',
            boxSizing: 'border-box',
          }}
        />
        )}

        {/* The header pane: follows the body sideways, never moves up or
            down. Its right padding is the width of the body's vertical
            scrollbar, so the two show the same span of days rather than the
            header running a scrollbar's width further. It is also where the
            canvas is dragged from — the grid itself draws tasks now. */}
        <div
          ref={panes.headerRef}
          title="Drag to move through time"
          style={{
            gridColumn: isMobile ? 1 : 2,
            gridRow: 1,
            overflow: 'hidden',
            background: 'var(--gantt-surface)',
            borderBottom: '1px solid var(--gantt-rule-strong)',
            boxSizing: 'border-box',
            paddingRight: panes.gutter.vertical,
            // The ruler is what the plan is pulled along by now — the canvas
            // under it draws a task instead. See useScrollPanes.
            cursor: 'grab',
            touchAction: 'none',
          }}
        >
          <TimelineHeader cells={headerCells} columnWidth={columnWidth} width={canvasWidth} />
        </div>

        {/* The task column, in whichever of its two shapes this width calls
            for.

            Above the breakpoint it is a pane in the grid: follows the body
            down, never moves sideways, and its bottom padding matches the
            body's horizontal scrollbar for the same reason the header's right
            padding matches the vertical one.

            Below it the same box leaves the grid and lies over the canvas as a
            drawer, anchored to the left edge — the side the column is on at
            every other width, so a name never changes which side of its bar it
            is read from. It spans the full height and gives its first 56px to
            a head of its own, which is exactly the corner block's height, so
            the rows underneath still start level with the bars.

            The wrapper and the pane inside it are one element each in both
            layouts, not two subtrees swapped at the breakpoint: the pane is
            what `panes.listRef` points at and what carries the wheel and
            scroll listeners, and a window dragged across 768px must not leave
            those on a node that is no longer in the document. */}
        <div
          id="gantt-task-drawer"
          ref={drawerRef}
          className={isMobile ? 'gantt-drawer' : undefined}
          data-open={isMobile ? isDrawerOpen : undefined}
          aria-label={isMobile ? 'Tasks' : undefined}
          onPointerDown={isMobile ? beginDrawerSwipe : undefined}
          style={{
            background: 'var(--gantt-surface)',
            borderRight: '1px solid var(--gantt-rule-strong)',
            boxSizing: 'border-box',
            overflow: 'hidden',
            ...(isMobile
              ? {
                  position: 'absolute',
                  top: 0,
                  bottom: 0,
                  left: 0,
                  width: 'var(--gantt-drawer-width)',
                  zIndex: 30,
                  display: 'flex',
                  flexDirection: 'column',
                  // The browser keeps the vertical axis and scrolls the list
                  // with it; the horizontal one is the swipe's.
                  touchAction: 'pan-y',
                }
              : { gridColumn: 1, gridRow: 2 }),
          }}
        >
          {isMobile && (
            <div
              style={{
                flex: 'none',
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                height: HEADER_HEIGHT_PX,
                padding: '0 6px 0 16px',
                borderBottom: '1px solid var(--gantt-rule-strong)',
                boxSizing: 'border-box',
              }}
            >
              <span style={{ flex: 1, minWidth: 0, fontSize: 15, fontWeight: 600, color: 'var(--gantt-text)' }}>
                Tasks
              </span>
              {/* The third way out, beside the swipe and the scrim. The header's
                  own Tasks button is the fourth and closes it too — it is a
                  toggle, not a one-way door. */}
              <button
                type="button"
                onClick={() => setDrawerOpen(false)}
                title="Close"
                aria-label="Close the task list"
                className={buttonBaseClass('ghost', 'h-11 w-11 flex-none text-muted-foreground')}
              >
                <X size={18} strokeWidth={1.8} aria-hidden="true" />
              </button>
            </div>
          )}

          <div
            ref={panes.listRef}
            style={{
              // Above the breakpoint the pane fills the grid cell and is the
              // hidden box the sync writes into; below it, it is the drawer's
              // scrolling body and takes the finger itself.
              ...(isMobile
                ? { flex: 1, minHeight: 0, overflowX: 'hidden', overflowY: 'auto', touchAction: 'pan-y' }
                : { height: '100%', overflow: 'hidden' }),
              paddingBottom: panes.gutter.horizontal,
            }}
          >
            <TaskList
              rows={rows}
              collapsed={collapsed}
              selectedId={selectedId}
              minHeight={bodyHeight}
              onSelect={(id) => {
                // A swipe that happened to start over a row is not a tap on
                // it. Same rule, same flag as a bar that was dragged.
                if (isMobile && movedRef.current) return;
                select(id);
                // The panel it opens is full-screen at this width, so leaving
                // the drawer standing behind it would only mean finding it
                // still there on the way out.
                if (isMobile) setDrawerOpen(false);
              }}
              onToggleCollapse={toggleCollapsed}
              onCycleStatus={cycleStatus}
              onRename={(id, name) => updateItem(id, { label: name })}
              onAddTask={addTask}
              onAddSubtask={addSubtask}
              onContextMenu={openMenu}
            />
          </div>
        </div>

        {/* Everything that is not the drawer, dimmed and inert while it is
            out — the chart, the ruler above it, and the strip of both showing
            past the drawer's right edge. A press anywhere on it puts the
            drawer away, which is the largest and most forgiving of the four
            ways to close it. The app's own header is outside this box and
            stays live: import and export are not things the drawer is in the
            middle of. */}
        {isMobile && (
          <div
            className="gantt-scrim"
            data-open={isDrawerOpen}
            onClick={() => setDrawerOpen(false)}
            aria-hidden="true"
          />
        )}

        {/* The body: the only scroller on the screen. Both bars belong to it,
            which puts the horizontal one along the bottom of the timeline and
            nowhere near the task names. `overflow-x: scroll` rather than
            `auto` so the bar is a standing part of the zone instead of
            something that appears and disappears under the pointer. */}
        <div
          ref={panes.bodyRef}
          className="gantt-scroll"
          style={{
            gridColumn: isMobile ? 1 : 2,
            gridRow: 2,
            overflowX: 'scroll',
            overflowY: 'auto',
            position: 'relative',
          }}
        >
          <TimelineBody
            rows={rows}
            spans={spans}
            scale={scale}
            cells={headerCells}
            barStyleById={barStyleById}
            columnWidth={columnWidth}
            width={canvasWidth}
            height={bodyHeight}
            todayIndex={todayIndex}
            weekendStarts={weekends}
            dateRangeById={dateRangeById}
            statusLabelById={statusLabelById}
            selectedId={selectedId}
            drag={drag}
            dragLabel={dragLabel}
            dayCount={renderedDays}
            scrollerRef={panes.bodyRef}
            isMobile={isMobile}
            formatDay={(index) => formatDayLabel(minDate, index)}
            onCreateTask={createTaskOnLane}
            onPointerDownBar={beginDrag}
            onContextMenuBar={openMenu}
            onSelectBar={(id) => {
              // A bar drag ends in a click too, and so does a pan that happens
              // to be released over a bar — opening the panel because someone
              // moved a bar, or scrolled past one, would be a surprise.
              if (!movedRef.current && !panes.isPanningRef.current) select(id);
            }}
          />
        </div>
      </div>

      {menuItem && menu && (
        <ContextMenu
          x={menu.x}
          y={menu.y}
          label={menuItem.label}
          actions={menuActions(menuItem.id)}
          onClose={() => setMenu(null)}
        />
      )}

      {selectedItem && (
        <EditTaskPanel
          // Keyed on the item, so switching rows gives the two date fields a
          // fresh start — each holds a draft of its own while a date is being
          // typed (see DateField), and that one is about the keystrokes, not
          // about the task. The half-written *comment* is no longer among
          // them: it is keyed by task id in the view store, so it survives the
          // panel closing and comes back when the task is opened again.
          key={selectedItem.id}
          item={selectedItem}
          minDate={minDate}
          spans={spans}
        />
      )}
    </div>
  );
}
