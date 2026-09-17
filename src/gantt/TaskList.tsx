import { useState } from 'react';
import { Plus } from 'lucide-react';
import { ADD_ROW_HEIGHT_PX } from './geometry';
import type { GanttRowModel } from './rows';
import { TaskListRow } from './TaskListRow';
import { useGanttViewStore } from './viewStore';

interface TaskListProps {
  rows: GanttRowModel[];
  collapsed: Record<string, boolean>;
  selectedId: string | null;
  onSelect: (id: string) => void;
  onToggleCollapse: (id: string) => void;
  onCycleStatus: (id: string) => void;
  onRename: (id: string, name: string) => void;
  onAddTask: (name: string) => void;
  /** Copies a row's whole branch into a plan of its own and opens it — what
   * the sub-task count badge does. */
  /** Creates a sub-task under `parentId`. The screen opens the new row's name
   * for editing as it goes, so the task is named where it will live rather
   * than in a form somewhere else. */
  onAddSubtask: (parentId: string) => void;
  /** A right-click on a row, with the pointer's viewport position. */
  onContextMenu: (id: string, event: React.MouseEvent) => void;
  /** The body's content height. The two panes scroll on one offset, so the
   * list has to be exactly as tall as the canvas beside it — otherwise the
   * offset that puts the last bar at the foot of the timeline puts the last
   * name somewhere else. */
  minHeight: number;
}

/** The task column.
 *
 * It fills its pane rather than naming a width of its own — the width is the
 * grid column's, which a drag on the seam can change (see `listWidth`).
 *
 * Its pane sits to the left of the plan and never scrolls sideways; the
 * shell writes its vertical offset from the body's as the rows scroll (see
 * useScrollPanes), which is what keeps a name level with its bar.
 *
 * Renaming and adding are the two pieces of text this column collects, and
 * both live here as local state rather than in the view store: they are a
 * keystroke's worth of draft, discarded on Escape, and nothing outside this
 * column has any use for either. */
export function TaskList({
  rows,
  collapsed,
  selectedId,
  onSelect,
  onToggleCollapse,
  onCycleStatus,
  onRename,
  onAddTask,
  onAddSubtask,
  onContextMenu,
  minHeight,
}: TaskListProps) {
  const renamingId = useGanttViewStore((state) => state.renamingId);
  const renameDraft = useGanttViewStore((state) => state.renameDraft);
  const beginRename = useGanttViewStore((state) => state.beginRename);
  const setRenameDraft = useGanttViewStore((state) => state.setRenameDraft);
  const endRename = useGanttViewStore((state) => state.endRename);

  const [isAdding, setIsAdding] = useState(false);
  const [addText, setAddText] = useState('');

  const commitEdit = () => {
    if (renamingId === null) return;
    const trimmed = renameDraft.trim();
    if (trimmed !== '') onRename(renamingId, trimmed);
    endRename();
  };

  const commitAdd = () => {
    const trimmed = addText.trim();
    if (trimmed === '') {
      setIsAdding(false);
      setAddText('');
      return;
    }
    onAddTask(trimmed);
    // The field stays open and empty: adding tasks is usually adding several,
    // and closing after each one would cost a click per task.
    setAddText('');
  };

  return (
    <div
      style={{
        width: '100%',
        minHeight,
        background: 'var(--gantt-surface)',
        boxSizing: 'border-box',
      }}
    >
      {rows.map((row) => (
        <TaskListRow
          key={row.item.id}
          row={row}
          isSelected={selectedId === row.item.id}
          isCollapsed={collapsed[row.item.id] === true}
          isEditing={renamingId === row.item.id}
          editText={renameDraft}
          onEditTextChange={setRenameDraft}
          onCommitEdit={commitEdit}
          onCancelEdit={endRename}
          onBeginEdit={() => beginRename(row.item.id, row.item.label)}
          onSelect={() => onSelect(row.item.id)}
          onToggleCollapse={() => onToggleCollapse(row.item.id)}
          onCycleStatus={() => onCycleStatus(row.item.id)}
          onAddSubtask={() => onAddSubtask(row.item.id)}
          onContextMenu={(event) => onContextMenu(row.item.id, event)}
        />
      ))}

      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 8,
          height: ADD_ROW_HEIGHT_PX,
          padding: '0 12px 0 44px',
          borderTop: '1px solid var(--gantt-rule-soft)',
          background: 'var(--gantt-add-row)',
        }}
      >
        {isAdding ? (
          <input
            type="text"
            autoFocus
            placeholder="Task name, then Enter"
            value={addText}
            onChange={(event) => setAddText(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') commitAdd();
              if (event.key === 'Escape') {
                setIsAdding(false);
                setAddText('');
              }
            }}
            onBlur={() => {
              setIsAdding(false);
              setAddText('');
            }}
            aria-label="New task name"
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
              color: 'var(--gantt-text)',
            }}
          />
        ) : (
          <>
            <button
              type="button"
              onClick={() => setIsAdding(true)}
              title="Add task"
              aria-label="Add task"
              className="gantt-add-button"
              style={{
                flex: 'none',
                border: 'none',
                borderRadius: 999,
                background: 'var(--gantt-add-button-bg)',
                color: 'var(--gantt-add-button-fg)',
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
              }}
            >
              <Plus size={16} strokeWidth={2} aria-hidden="true" />
            </button>
            {/* One line, ellipsised: the row is a fixed 46px and a second
                line would spill out of it. */}
            <span
              title="Add task · double-click a name to rename"
              style={{
                fontSize: 12,
                color: 'var(--gantt-text-muted)',
                minWidth: 0,
                whiteSpace: 'nowrap',
                overflow: 'hidden',
                textOverflow: 'ellipsis',
              }}
            >
              Add task · double-click a name to rename
            </span>
          </>
        )}
      </div>
    </div>
  );
}
