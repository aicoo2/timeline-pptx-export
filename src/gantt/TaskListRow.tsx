import { ChevronDown, ChevronRight, Layers, Plus } from 'lucide-react';
import { rowPaddingLeft, ROW_HEIGHT_PX } from './geometry';
import type { GanttRowModel } from './rows';
import { STATUS_LABEL } from './tone';
import { StatusIcon } from './StatusIcon';

interface TaskListRowProps {
  row: GanttRowModel;
  isSelected: boolean;
  isCollapsed: boolean;
  isEditing: boolean;
  editText: string;
  onEditTextChange: (value: string) => void;
  onCommitEdit: () => void;
  onCancelEdit: () => void;
  onBeginEdit: () => void;
  onSelect: () => void;
  onToggleCollapse: () => void;
  onCycleStatus: () => void;
  /** Lift this row and everything under it out into a plan of its own. Only
   * ever called from a group's count badge, which is the one row element that
   * names the sub-tasks it would take with it. */
  /** Create a sub-task under this row and start renaming it. */
  onAddSubtask: () => void;
  onContextMenu: (event: React.MouseEvent) => void;
}

/** The drag-to-reorder grip: a 2×3 grid of 3px dots.
 *
 * Drawn and not wired, exactly as in the handoff, whose own README lists
 * "row reordering by the drag-dots handle is drawn but not wired" as one of
 * its three known gaps. It keeps the row's spacing and the affordance the
 * design shows; what it does when dragged is a question the handoff leaves
 * open rather than one to answer here. It is also the first thing to go below
 * the mobile breakpoint (see .gantt-row-grip): a row 300px wide has no space
 * to spend advertising a gesture that does nothing. */
function DragGrip() {
  return (
    <span
      className="gantt-row-grip"
      title="Drag to reorder"
      // `display` is in the stylesheet, not here: an inline one would outrank
      // the media query that takes this off the row below the breakpoint.
      style={{ gridTemplateColumns: '3px 3px', gap: 3, flex: 'none', cursor: 'grab', opacity: 0.85 }}
    >
      {Array.from({ length: 6 }, (_, index) => (
        <span
          key={index}
          style={{ width: 3, height: 3, borderRadius: 9, background: 'var(--gantt-text-faint)' }}
        />
      ))}
    </span>
  );
}

/** One line of the task list: what the row is, and what can be done to it.
 *
 * A group starts hard against the column's edge because its caret occupies
 * that space; a task starts where the caret would have ended. Both names are
 * plain text; what separates a top-level task from its work is weight and
 * size, not a frame around one of them. */
export function TaskListRow({
  row,
  isSelected,
  isCollapsed,
  isEditing,
  editText,
  onEditTextChange,
  onCommitEdit,
  onCancelEdit,
  onBeginEdit,
  onSelect,
  onToggleCollapse,
  onCycleStatus,
  onAddSubtask,
  onContextMenu,
}: TaskListRowProps) {
  const { item, depth, isGroup, childCount, isSubtask, status } = row;

  // What the weight and the colour say is "top level", not "has sub-tasks". A
  // root that has not been given any yet is still a root and can take them at
  // any moment, so it is set in the same semibold and the same full-strength
  // text colour as one that already has them, and a group nested under a root
  // is set plain and muted like the rest of its level.
  //
  // Weight and colour are the only things that read depth this way. The caret,
  // the count badge, the row's indent and the "(rolled up)" status all still
  // turn on having children, because each of them is about children that
  // exist; the bar's colour already went by depth in the whole plan and is
  // unchanged.
  const isRoot = depth === 0;

  // One признак, one pair of tokens, and the tokens are what carries the two
  // themes: `--gantt-text` and `--gantt-text-muted` swap to slate-300/500 in
  // the dark block of tokens.css, so the contrast between a root and its
  // sub-tasks survives the theme without either colour being written here.
  //
  // This replaces the older rule on tasks, where an untouched (`todo`) task
  // was muted and a started one was not. Status is already said twice on the
  // row — by the icon and by its title — and saying it a third time in the
  // name's colour was what left two roots looking like different levels.
  const nameColor = isRoot ? 'var(--gantt-text)' : 'var(--gantt-text-muted)';

  // Shared by both: one line, ellipsised, and a box for the hover wash to
  // fill. The wash is what is left of the pill — see .gantt-row-name.
  const nameBase: React.CSSProperties = {
    flex: 1,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    borderRadius: 6,
    boxSizing: 'border-box',
    fontWeight: isRoot ? 600 : 400,
  };

  const nameStyle: React.CSSProperties = isGroup
    ? {
        ...nameBase,
        fontSize: 15,
        color: nameColor,
        // A group's name carried no box of its own. It gets the task's, and
        // an equal negative margin on the left hands the space straight back,
        // so the name stays exactly where the handoff puts it and the wash
        // still has the same 11px of run-up before the first letter.
        border: '1px solid transparent',
        padding: '6px 10px',
        marginLeft: -11,
        lineHeight: '17px',
      }
    : {
        ...nameBase,
        fontSize: 14,
        color: nameColor,
        // The pill's border and fill are gone; its box is not. A transparent
        // border of the same width keeps every name on the pixel it was on,
        // and keeps the two name treatments one shape.
        border: '1px solid transparent',
        padding: '6px 10px',
        lineHeight: '17px',
      };

  return (
    <div
      onClick={onSelect}
      onContextMenu={onContextMenu}
      className="gantt-row"
      style={{
        display: 'flex',
        alignItems: 'center',
        gap: 10,
        height: ROW_HEIGHT_PX,
        boxSizing: 'border-box',
        padding: `0 10px 0 ${rowPaddingLeft(depth, isGroup)}px`,
        borderBottom: '1px solid var(--gantt-rule-soft)',
        cursor: 'pointer',
        background: isSelected ? 'var(--gantt-row-selected)' : 'var(--gantt-surface)',
      }}
    >
      <button
        type="button"
        onClick={(event) => {
          event.stopPropagation();
          onToggleCollapse();
        }}
        title="Show / hide sub-tasks"
        aria-label="Show / hide sub-tasks"
        aria-expanded={!isCollapsed}
        className="gantt-row-caret"
        style={{
          flex: 'none',
          border: 'none',
          background: 'transparent',
          padding: 0,
          cursor: 'pointer',
          display: isGroup ? 'flex' : 'none',
          alignItems: 'center',
          justifyContent: 'center',
        }}
      >
        {isCollapsed ? (
          <ChevronRight size={14} strokeWidth={2.2} color="var(--gantt-text-secondary)" />
        ) : (
          <ChevronDown size={14} strokeWidth={2.2} color="var(--gantt-text-secondary)" />
        )}
      </button>

      <DragGrip />

      <button
        type="button"
        onClick={(event) => {
          event.stopPropagation();
          onCycleStatus();
        }}
        // A group's status is its children's, so there is nothing here to
        // change; the title says so instead of the control lying about it.
        title={`${STATUS_LABEL[status]}${isGroup ? ' (rolled up)' : ' — click to change'}`}
        aria-label={`${item.label}: ${STATUS_LABEL[status]}`}
        className="gantt-row-status"
        style={{
          flex: 'none',
          border: 'none',
          background: 'transparent',
          padding: 0,
          cursor: 'pointer',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
        }}
      >
        <StatusIcon status={status} />
      </button>

      {isEditing ? (
        <input
          type="text"
          value={editText}
          autoFocus
          // Selected, not just focused: a row that was created a moment ago
          // opens this field on a placeholder name, and typing has to replace
          // it rather than run on the end of it. A rename works the same way,
          // which is what an inline rename does everywhere else.
          onFocus={(event) => event.target.select()}
          onClick={(event) => event.stopPropagation()}
          onChange={(event) => onEditTextChange(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') onCommitEdit();
            if (event.key === 'Escape') onCancelEdit();
          }}
          onBlur={onCommitEdit}
          aria-label="Task name"
          className="gantt-field"
          style={{
            flex: 1,
            minWidth: 0,
            height: 30,
            boxSizing: 'border-box',
            border: '1px solid var(--gantt-edit-focus)',
            borderRadius: 6,
            padding: '0 10px',
            outline: 'none',
            background: 'var(--gantt-edit-bg)',
            color: 'var(--gantt-text)',
          }}
        />
      ) : (
        <span
          className="gantt-row-name"
          onDoubleClick={(event) => {
            event.stopPropagation();
            onBeginEdit();
          }}
          // The name, and then what can be done to it — the form the status
          // button beside it already uses ("In progress — click to change").
          // Every other element on this row names its own action in its
          // title; the name was the one that only repeated itself, and since
          // the pill came off it is also the one element with nothing at rest
          // to say it is a field. The name stays first, because the other job
          // this title has is showing a long name the column had to ellipsise.
          title={`${item.label} — double-click to rename`}
          style={nameStyle}
        >
          {item.label}
        </span>
      )}

      {/* How many sub-tasks the row has, and nothing more. It used to carry a
          click that copied the branch into a plan of its own and opened it,
          which is the one thing this badge must not do: a plan is something
          made deliberately from "New plan", and a branch opened as one both
          hid the rest of the plan and left a second entry in the switcher
          that nobody had asked for. A count is a fact about the row, so it is
          rendered as one. */}
      {isGroup && (
        <span
          className="gantt-subcount"
          title={`${childCount} sub-task${childCount === 1 ? '' : 's'}`}
          style={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: 3,
            flex: 'none',
            minWidth: 18,
            height: 17,
            padding: '0 6px',
            border: 'none',
            borderRadius: 999,
            background: 'var(--gantt-subcount-bg)',
            color: 'var(--gantt-surface)',
            fontSize: 10,
            fontWeight: 600,
          }}
        >
          <Layers size={10} strokeWidth={2.4} aria-hidden="true" />
          {childCount}
        </span>
      )}

      {/* A top-level row can take a sub-task — a task with one is simply a
          group from then on — and a sub-task cannot: the plan is built one
          level deep, so the control is absent rather than present and
          refusing. Out of the way until the pointer is on the row: at rest
          the column is names, not controls.

          The slot stays behind it, empty. The button is invisible at rest
          already (see .gantt-row-add), so taking its 18px away as well would
          widen a sub-task's name pill past its parent's and leave the column
          with two right edges — for a control the eye never sees anyway. */}
      {isSubtask ? (
        <span aria-hidden="true" className="gantt-row-add-slot" style={{ flex: 'none' }} />
      ) : (
        <button
          type="button"
          className="gantt-row-add"
          onClick={(event) => {
            event.stopPropagation();
            onAddSubtask();
          }}
          title="Add sub-task"
          aria-label={`Add a sub-task under ${item.label}`}
          style={{
            flex: 'none',
            border: 'none',
            borderRadius: 4,
            background: 'transparent',
            padding: 0,
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
          }}
        >
          <Plus size={14} strokeWidth={2.2} aria-hidden="true" />
        </button>
      )}
    </div>
  );
}
