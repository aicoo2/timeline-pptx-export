/** Naming a plan, when the library may already hold the name asked for.
 *
 * This file used to do more: it also copied a branch — a task and its whole
 * sub-tree — into a plan of its own, for the sub-task count badge's click.
 * That action is gone (a plan is made from "New plan" and nothing else), and
 * with it `copyBranch`, its `BranchCopy` shape and the notice that reported
 * the dependency links such a copy had to drop. What is left is the one part
 * that was never about branches. */

/** What to call a plan when the library may already hold the name asked for
 * — on creating one, and on renaming one.
 *
 * The plain label wherever it is free, and the label with a number after it
 * where it isn't — the same thing a file manager does with a second copy. The
 * switcher lists plans by name alone, so two plans called "Design" would be
 * two rows nobody can tell apart. */
export function uniquePlanName(base: string, taken: string[]): string {
  const existing = new Set(taken);
  if (!existing.has(base)) return base;

  let suffix = 2;
  while (existing.has(`${base} ${suffix}`)) suffix += 1;
  return `${base} ${suffix}`;
}
