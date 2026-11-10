import { beforeEach, describe, expect, it } from 'vitest';
import { readOneOfficeTaskPageState } from '../../src/oneoffice/projectBatchCollector';

beforeEach(() => {
  window.history.replaceState({}, '', '/apps/work-task-task/view?ID=1001');
  document.body.innerHTML = `
    <div class="cards comments" comments-id="work-task-task@1001">
      <div class="cards-body"></div>
    </div>`;
});

describe('readOneOfficeTaskPageState', () => {
  it('keeps normal-work and project-task discussion namespaces separate', () => {
    window.history.replaceState({}, '', '/apps/work-normal-normal/view?ID=1001');
    const options = { allowedOrigin: window.location.origin, taskId: '1001', taskPath: '/apps/work-normal-normal/view' };
    expect(readOneOfficeTaskPageState(options).ready).toBe(false);
    document.body.innerHTML = '<div class="cards comments" comments-id="work-normal-normal@1001"><div class="cards-body"></div></div>';
    expect(readOneOfficeTaskPageState(options).ready).toBe(true);
    expect(readOneOfficeTaskPageState({ ...options, taskPath: undefined }).ready).toBe(false);
  });
  it('accepts the full task page when the matching discussion is visible', () => {
    expect(
      readOneOfficeTaskPageState({
        allowedOrigin: window.location.origin,
        taskId: '1001',
      }).ready,
    ).toBe(true);
  });

  it('rejects a different task ID', () => {
    expect(
      readOneOfficeTaskPageState({
        allowedOrigin: window.location.origin,
        taskId: '1002',
      }).ready,
    ).toBe(false);
  });

  it('counts root comments and replies in the rendered discussion', () => {
    document.querySelector('.cards-body')?.insertAdjacentHTML(
      'beforeend',
      `
        <div class="comment comment3001">
          <div class="comment-contents">
            <div class="cards sub-comments" comments-id="comment@3001">
              <div class="cards-body">
                <div class="comment comment3002"></div>
              </div>
            </div>
          </div>
        </div>
      `,
    );

    expect(
      readOneOfficeTaskPageState({
        allowedOrigin: window.location.origin,
        taskId: '1001',
      }),
    ).toMatchObject({
      ready: true,
      rootCount: 1,
      replyCount: 1,
    });
  });

  it('rejects hidden or incomplete discussion DOM', () => {
    document.querySelector('.cards.comments')?.classList.add('hidden');

    expect(
      readOneOfficeTaskPageState({
        allowedOrigin: window.location.origin,
        taskId: '1001',
      }).ready,
    ).toBe(false);
  });
});
