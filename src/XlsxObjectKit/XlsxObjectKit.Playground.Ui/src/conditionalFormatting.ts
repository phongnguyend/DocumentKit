import { parse, type Cell, type ConditionalRule, type Style } from './types';

export function conditionalStyle(cell: Cell | undefined, style: Style, rules: ConditionalRule[]): Style {
  if (!cell || (cell.formula != null && !cell.hasCachedValue)) {
    return style;
  }
  const number = Number(cell.value);

  const [row, column] = parse(cell.address);
  const changes: Partial<Style> = {};
  for (const rule of rules) {
    if (!rule.supported || rule.value == null && rule.matchedCells == null) {
      continue;
    }
    const inside = rule.range.split(/\s+/).some(area => {
      try {
        const [first, last = first] = area.replaceAll('$', '').split(':').map(parse);
        return row >= first[0] && row <= last[0] && column >= first[1] && column <= last[1];
      } catch {
        return false;
      }
    });
    const a = rule.value ?? 0;
    const b = rule.secondValue ?? a;
    const match = rule.matchedCells != null ? rule.matchedCells.includes(cell.address) : cell.type === 'n' && cell.value.trim() !== '' && Number.isFinite(number) && ({ greaterThan: number > a, greaterThanOrEqual: number >= a, lessThan: number < a,
      lessThanOrEqual: number <= a, equal: number === a, notEqual: number !== a,
      between: number >= a && number <= b, notBetween: number < a || number > b } as Record<string, boolean>)[rule.operator];
    if (inside && match) {
      changes.bold ??= rule.bold ?? undefined;
      changes.italic ??= rule.italic ?? undefined;
      changes.color ??= rule.color ?? undefined;
      changes.background ??= rule.background ?? undefined;
      if (rule.stopIfTrue) {
        break;
      }
    }
  }
  const applied = Object.fromEntries(Object.entries(changes).filter(([, value]) => value != null));
  return Object.keys(applied).length === 0 ? style : { ...style, ...applied };
}
