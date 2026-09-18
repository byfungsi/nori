/** Directed dependency index. Keys are opaque to the formula package. */
export class DependencyGraph {
  private readonly precedents = new Map<string, ReadonlySet<string>>();
  private readonly dependents = new Map<string, Set<string>>();
  /** Replace all outgoing dependencies and remove obsolete reverse edges. */
  set(key: string, dependencies: Iterable<string>): void {
    this.remove(key);
    const values = new Set(dependencies);
    this.precedents.set(key, values);
    for (const value of values) {
      const reverse = this.dependents.get(value) ?? new Set<string>();
      reverse.add(key);
      this.dependents.set(value, reverse);
    }
  }
  /** Remove a formula's outgoing edges while preserving cells that depend on it. */
  remove(key: string): void {
    for (const value of this.precedents.get(key) ?? []) {
      const reverse = this.dependents.get(value);
      reverse?.delete(key);
      if (!reverse?.size) this.dependents.delete(value);
    }
    this.precedents.delete(key);
  }
  /** Return transitive dependents, including a cycle origin if it reaches itself. */
  affectedBy(key: string): ReadonlySet<string> {
    const found = new Set<string>();
    const queue = [key];
    while (queue.length) {
      const item = queue.pop();
      if (item === undefined) break;
      for (const next of this.dependents.get(item) ?? [])
        if (!found.has(next)) {
          found.add(next);
          queue.push(next);
        }
    }
    return found;
  }
}
