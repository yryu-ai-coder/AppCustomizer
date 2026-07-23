import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

const LOG_SOURCE: string = 'SiteLogoRedirectApplicationCustomizer';
const TOAST_STYLE_ID: string = 'spfx-redirect-toast-styles';
const REDIRECT_DELAY_MS: number = 2000;

export interface ISiteLogoRedirectApplicationCustomizerProperties {
  redirectUrl: string;
  openInNewTab?: boolean;
}

// Containers that hold the modern SharePoint site header (logo + title area).
// NOTE: this must NOT include #SuiteNavWrapper/#O365_NavHeader - that's the
// tenant-level O365 suite bar (waffle + tenant logo), a different element that
// happens to link to the same URL when the site is the root site collection.
const HEADER_CONTAINER_SELECTOR: string =
  '#spSiteHeader, [data-automationid="SiteHeader"], header[role="banner"]';

// The O365 suite bar (tenant logo, app launcher, search, etc.) - always excluded.
const SUITE_BAR_SELECTOR: string =
  '#SuiteNavWrapper, #O365_NavHeader, [id^="O365_"]';

// Known data-automationid/class values used by SharePoint for the logo/title elements.
// SharePoint's header markup changes across UI updates, so we match several known patterns.
const LOGO_TITLE_SELECTOR: string = [
  '[data-automationid="SiteHeaderTitle"]',
  '[data-automationid="SiteHeaderLogo"]',
  '[data-automationid="SiteLogoContainer"]',
  '[data-automationid="SiteLogoImageOnly"]',
  '.ms-siteLogo-actionButton',
  '.ms-siteLogoContainer'
].join(', ');

export default class SiteLogoRedirectApplicationCustomizer
  extends BaseApplicationCustomizer<ISiteLogoRedirectApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${LOG_SOURCE}`);

    if (!this.properties.redirectUrl) {
      Log.warn(LOG_SOURCE, 'redirectUrl property is not set; site logo/title click will not be intercepted.');
      return Promise.resolve();
    }

    document.addEventListener('click', this._onDocumentClick, true);

    return Promise.resolve();
  }

  public onDispose(): void {
    document.removeEventListener('click', this._onDocumentClick, true);
  }

  private _onDocumentClick = (event: MouseEvent): void => {
    const target: HTMLElement | null = event.target as HTMLElement;
    if (!target) {
      return;
    }

    if (target.closest(SUITE_BAR_SELECTOR)) {
      return;
    }

    const headerContainer: Element | null = target.closest(HEADER_CONTAINER_SELECTOR);
    if (!headerContainer) {
      return;
    }

    const matchedByAutomationId: boolean = !!target.closest(LOGO_TITLE_SELECTOR);
    const matchedByHomeLink: boolean = this._isSiteHomeAnchor(target);

    if (!matchedByAutomationId && !matchedByHomeLink) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    const redirectUrl: string = this.properties.redirectUrl;

    if (this.properties.openInNewTab) {
      // Open synchronously within the click handler so the popup blocker
      // still treats it as a user-initiated action.
      window.open(redirectUrl, '_blank', 'noopener');
      this._showRedirectToast(redirectUrl);
    } else {
      this._showRedirectToast(redirectUrl, () => {
        window.location.href = redirectUrl;
      });
    }
  }

  private _isSiteHomeAnchor(target: HTMLElement): boolean {
    const anchor: HTMLAnchorElement | null = target.closest('a[href]') as HTMLAnchorElement | null;
    if (!anchor) {
      return false;
    }

    const siteUrl: string = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const anchorUrl: string = anchor.href.replace(/\/$/, '');

    return anchorUrl === siteUrl;
  }

  private _showRedirectToast(url: string, onComplete?: () => void): void {
    this._ensureToastStyles();

    const toast: HTMLDivElement = document.createElement('div');
    toast.className = 'spfx-redirect-toast';
    toast.setAttribute('role', 'status');
    toast.innerHTML = `
      <div class="spfx-redirect-toast__accent"></div>
      <div class="spfx-redirect-toast__body">
        <div class="spfx-redirect-toast__icon">&#8599;</div>
        <div class="spfx-redirect-toast__text">
          <div class="spfx-redirect-toast__title">Redirecting to this url&hellip;</div>
          <div class="spfx-redirect-toast__url"></div>
        </div>
      </div>
      <div class="spfx-redirect-toast__progress"><div class="spfx-redirect-toast__progress-bar"></div></div>
    `;

    const urlEl: Element | null = toast.querySelector('.spfx-redirect-toast__url');
    if (urlEl) {
      urlEl.textContent = url;
    }

    document.body.appendChild(toast);

    // Kick off the entrance + progress-bar animations on the next frame.
    window.requestAnimationFrame(() => {
      toast.classList.add('spfx-redirect-toast--visible');
      const progressBar: HTMLElement | null = toast.querySelector('.spfx-redirect-toast__progress-bar');
      if (progressBar) {
        progressBar.style.transitionDuration = `${REDIRECT_DELAY_MS}ms`;
        progressBar.style.width = '0%';
      }
    });

    window.setTimeout(() => {
      toast.classList.add('spfx-redirect-toast--closing');
      window.setTimeout(() => toast.remove(), 200);
      if (onComplete) {
        onComplete();
      }
    }, REDIRECT_DELAY_MS);
  }

  private _ensureToastStyles(): void {
    if (document.getElementById(TOAST_STYLE_ID)) {
      return;
    }

    const style: HTMLStyleElement = document.createElement('style');
    style.id = TOAST_STYLE_ID;
    style.textContent = `
      .spfx-redirect-toast {
        position: fixed;
        top: 24px;
        right: 24px;
        z-index: 2147483000;
        min-width: 300px;
        max-width: 380px;
        background: #ffffff;
        border-radius: 12px;
        box-shadow: 0 10px 28px rgba(0, 0, 0, 0.16), 0 2px 8px rgba(0, 0, 0, 0.08);
        overflow: hidden;
        font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
        opacity: 0;
        transform: translateY(-12px);
        transition: opacity 0.25s ease-out, transform 0.25s ease-out;
      }
      .spfx-redirect-toast--visible {
        opacity: 1;
        transform: translateY(0);
      }
      .spfx-redirect-toast--closing {
        opacity: 0;
        transform: translateY(-8px);
      }
      .spfx-redirect-toast__accent {
        height: 4px;
        background: linear-gradient(90deg, #0078d4, #50e6ff);
      }
      .spfx-redirect-toast__body {
        display: flex;
        align-items: flex-start;
        gap: 12px;
        padding: 14px 16px 10px 16px;
      }
      .spfx-redirect-toast__icon {
        flex: 0 0 auto;
        width: 32px;
        height: 32px;
        border-radius: 50%;
        background: #eff6fc;
        color: #0078d4;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 16px;
      }
      .spfx-redirect-toast__text {
        min-width: 0;
      }
      .spfx-redirect-toast__title {
        font-size: 14px;
        font-weight: 600;
        color: #201f1e;
      }
      .spfx-redirect-toast__url {
        margin-top: 2px;
        font-size: 12px;
        color: #605e5c;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      .spfx-redirect-toast__progress {
        height: 3px;
        background: #f3f2f1;
      }
      .spfx-redirect-toast__progress-bar {
        height: 100%;
        width: 100%;
        background: #0078d4;
        transition: width linear;
      }
    `;
    document.head.appendChild(style);
  }
}
