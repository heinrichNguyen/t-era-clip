import { beforeEach, describe, expect, it } from 'vitest';
import projectFixture from '../fixtures/oneoffice/project-discussion.html?raw';
import {
  collectOneOfficeDiscussionFromPage,
  ONEOFFICE_PROGRESS_MESSAGE,
} from '../../src/oneoffice/collector';
import {
  isOneOfficeFilteredTaskListUrl,
  isOneOfficeProjectPageUrl,
  isSupportedOneOfficeUrl,
  parseOneOfficeSnapshot,
} from '../../src/oneoffice/collectionClient';

const collectorOptions = {
  allowedOrigin: window.location.origin,
  maxLoadMoreClicks: 5,
  noProgressTimeoutMs: 500,
  progressMessageType: ONEOFFICE_PROGRESS_MESSAGE,
  settleTimeMs: 20,
};

beforeEach(() => {
  document.body.innerHTML = '';
});

describe('1Office URL validation', () => {
  it('accepts only the approved HTTPS origin', () => {
    expect(isSupportedOneOfficeUrl('https://office.meygroup.vn/apps/work-project-project/view?ID=1')).toBe(
      true,
    );
    expect(isSupportedOneOfficeUrl('http://office.meygroup.vn/')).toBe(false);
    expect(isSupportedOneOfficeUrl('https://office.meygroup.vn.attacker.test/')).toBe(false);
    expect(isSupportedOneOfficeUrl('not-a-url')).toBe(false);
    expect(
      isOneOfficeProjectPageUrl(
        'https://office.meygroup.vn/apps/work-project-project/view?ID=830&tab=work',
      ),
    ).toBe(true);
    expect(
      isOneOfficeProjectPageUrl('https://office.meygroup.vn/apps/work-task-task/quickview?ID=830'),
    ).toBe(false);
    expect(
      isOneOfficeFilteredTaskListUrl(
        'https://office.meygroup.vn/apps/work-task-task/list?@sfilter=%7B%22ors%22%3A%5B%5D%7D',
      ),
    ).toBe(true);
    expect(
      isOneOfficeFilteredTaskListUrl(
        'https://office.meygroup.vn/apps/work-task-task/view?ID=1001',
      ),
    ).toBe(false);
  });
});

describe('collectOneOfficeDiscussionFromPage', () => {
  it('returns a minimized snapshot that can be parsed by the Gate 1 parser', async () => {
    document.body.innerHTML = `${projectFixture}
      <div style="display: none">
        <div class="detail-field">
          <div class="detail-field-label">Tên dự án</div>
          <div class="detail-field-content">Hidden stale project</div>
        </div>
      </div>
      <form><input value="not collected"></form>`;

    const snapshot = await collectOneOfficeDiscussionFromPage(collectorOptions);
    const parsed = parseOneOfficeSnapshot(snapshot);

    expect(snapshot.complete).toBe(true);
    expect(snapshot.loadMoreClicks).toBe(0);
    expect(snapshot.discussionHtml).not.toContain('<form');
    expect(snapshot.detailFieldsHtml).toContain('Tên dự án');
    expect(snapshot.detailFieldsHtml).not.toContain('Hidden stale project');
    expect(parsed.collection.loadedRootCommentCount).toBe(27);
    expect(parsed.collection.loadedReplyCount).toBe(5);
  });

  it('activates load-more until the control disappears', async () => {
    document.body.innerHTML = `
      <div class="cards comments" comments-id="work-project-project@800">
        <div class="prev"><a class="bt">Xem thêm 1/1 thảo luận cũ hơn</a></div>
        <div class="cards-body">
          <div class="comment comment5001"></div>
        </div>
      </div>`;
    const root = document.querySelector<HTMLElement>('.cards.comments');
    const button = document.querySelector<HTMLElement>('.prev .bt');

    button?.addEventListener('click', () => {
      window.setTimeout(() => {
        root?.querySelector('.cards-body')?.insertAdjacentHTML(
          'afterbegin',
          '<div class="comment comment5000"></div>',
        );
        root?.querySelector('.prev')?.remove();
      }, 10);
    });

    const snapshot = await collectOneOfficeDiscussionFromPage(collectorOptions);

    expect(snapshot.complete).toBe(true);
    expect(snapshot.loadMoreClicks).toBe(1);
    expect(snapshot.expectedRootCommentCount).toBe(2);
    expect(snapshot.discussionHtml).toContain('comment5000');
  });

  it('stops with an incomplete warning when load-more makes no progress', async () => {
    document.body.innerHTML = `
      <div class="cards comments" comments-id="work-task-task@801">
        <div class="prev"><a class="bt">Xem thêm 1/1 thảo luận cũ hơn</a></div>
        <div class="cards-body"><div class="comment comment6001"></div></div>
      </div>`;

    const snapshot = await collectOneOfficeDiscussionFromPage({
      ...collectorOptions,
      noProgressTimeoutMs: 20,
    });

    expect(snapshot.complete).toBe(false);
    expect(snapshot.loadMoreClicks).toBe(1);
    expect(snapshot.warnings[0]).toContain('did not change');
  });

  it('counts repeated root nodes only once when they share a comment ID', async () => {
    document.body.innerHTML = `
      <div class="cards comments" comments-id="work-project-project@802">
        <div class="cards-body">
          <div class="comment comment7001"></div>
          <div class="comment comment7001"></div>
          <div class="comment comment7002"></div>
        </div>
      </div>`;

    const snapshot = await collectOneOfficeDiscussionFromPage(collectorOptions);

    expect(snapshot.expectedRootCommentCount).toBe(2);
  });

  it('waits for the expected final root count after load-more disappears', async () => {
    document.body.innerHTML = `
      <div class="cards comments" comments-id="work-project-project@803">
        <div class="prev"><a class="bt">Xem thêm 2/2 thảo luận cũ hơn</a></div>
        <div class="cards-body"><div class="comment comment8001"></div></div>
      </div>`;
    const root = document.querySelector<HTMLElement>('.cards.comments');
    const body = root?.querySelector<HTMLElement>('.cards-body');
    const button = root?.querySelector<HTMLElement>('.prev .bt');

    button?.addEventListener('click', () => {
      root?.querySelector('.prev')?.remove();
      window.setTimeout(() => {
        body?.insertAdjacentHTML('beforeend', '<div class="comment comment8002"></div>');
      }, 10);
      window.setTimeout(() => {
        body?.insertAdjacentHTML('beforeend', '<div class="comment comment8003"></div>');
      }, 80);
    });

    const snapshot = await collectOneOfficeDiscussionFromPage(collectorOptions);

    expect(snapshot.complete).toBe(true);
    expect(snapshot.expectedRootCommentCount).toBe(3);
    expect(snapshot.discussionHtml).toContain('comment8003');
  });

  it('reports incomplete when load-more disappears before the expected count arrives', async () => {
    document.body.innerHTML = `
      <div class="cards comments" comments-id="work-project-project@804">
        <div class="prev"><a class="bt">Xem thêm 1/1 thảo luận cũ hơn</a></div>
        <div class="cards-body"><div class="comment comment9001"></div></div>
      </div>`;
    const button = document.querySelector<HTMLElement>('.prev .bt');

    button?.addEventListener('click', () => {
      document.querySelector('.prev')?.remove();
    });

    const snapshot = await collectOneOfficeDiscussionFromPage({
      ...collectorOptions,
      noProgressTimeoutMs: 50,
    });

    expect(snapshot.complete).toBe(false);
    expect(snapshot.warnings).toContain('Collection stopped at 1 of 2 expected root comments.');
  });

  it('keeps the expected total from the first load-more label', async () => {
    document.body.innerHTML = `
      <div class="cards comments" comments-id="work-project-project@805">
        <div class="prev"><a class="bt">Xem thêm 1/2 thảo luận cũ hơn</a></div>
        <div class="cards-body"><div class="comment comment9101"></div></div>
      </div>`;
    const root = document.querySelector<HTMLElement>('.cards.comments');
    const body = root?.querySelector<HTMLElement>('.cards-body');
    const button = root?.querySelector<HTMLElement>('.prev .bt');
    let clickCount = 0;

    button?.addEventListener('click', () => {
      clickCount += 1;
      if (clickCount === 1) {
        body?.insertAdjacentHTML('beforeend', '<div class="comment comment9102"></div>');
        if (button) {
          button.textContent = 'Xem thêm 1/5 thảo luận cũ hơn';
        }
        return;
      }

      body?.insertAdjacentHTML('beforeend', '<div class="comment comment9103"></div>');
      root?.querySelector('.prev')?.remove();
    });

    const snapshot = await collectOneOfficeDiscussionFromPage(collectorOptions);

    expect(snapshot.complete).toBe(true);
    expect(snapshot.expectedRootCommentCount).toBe(3);
  });
});
