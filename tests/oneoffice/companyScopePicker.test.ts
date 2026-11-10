// @vitest-environment jsdom
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { CompanyScopePicker } from '../../src/access-review/CompanyScopePicker';

it('shows group cards and a single member selector instead of a long row of company buttons', async () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
  const container = document.createElement('div'); document.body.append(container);
  const root = createRoot(container);
  const onChange = vi.fn();
  try {
    await act(async () => root.render(React.createElement(CompanyScopePicker, { value: 'TDK', disabled: false, onChange })));
    expect(container.querySelectorAll('button')).toHaveLength(6);
    let select = container.querySelector('select')!;
    expect([...select.options].map((item) => item.value)).toEqual(['TDK', 'TDH', 'TDP', 'TDC', 'TDW']);
    await act(async () => { select.value = 'TDH'; select.dispatchEvent(new Event('change', { bubbles: true })); });
    expect(onChange).toHaveBeenLastCalledWith('TDH');
    await act(async () => root.render(React.createElement(CompanyScopePicker, { value: 'TDU.CKKG', disabled: false, onChange })));
    select = container.querySelector('select')!;
    expect(select.value).toBe('TDU.CKKG');
    expect([...select.options].map((item) => item.value)).toEqual(['TDU', 'TDU.CKKG', 'TDU.GNKG', 'TDU.VLXD', 'TDU.XMHT']);
    const tdg = container.querySelector<HTMLButtonElement>('[aria-label="Chọn công ty TDG"]')!;
    await act(async () => tdg.click()); expect(onChange).toHaveBeenLastCalledWith('TDG');
    await act(async () => root.render(React.createElement(CompanyScopePicker, { value: 'TDG', disabled: false, onChange })));
    expect(container.querySelector('select')).toBeNull();
    expect(container.querySelector('[aria-label="Chọn công ty TDG"]')?.getAttribute('aria-pressed')).toBe('true');
    await act(async () => root.render(React.createElement(CompanyScopePicker, { value: 'TDG', disabled: true, onChange })));
    expect(container.querySelector('fieldset')?.disabled).toBe(true);
  } finally { await act(async () => root.unmount()); container.remove(); vi.unstubAllGlobals(); }
});
