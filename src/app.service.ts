import { Injectable } from '@nestjs/common';
import { Browser, Page } from 'puppeteer';
import { getBrowser } from './puppeteer.utils';
import { of } from 'rxjs';

const ICON_PATH = '/assets/icon-slack.png';
const CLIENT_URL = 'https://outlook.office.com/mail';

@Injectable()
export class AppService {
  browser: Browser;
  page: Page;
  initialized = false;

  async initialize() {
    this.browser = await getBrowser();
    this.page = await this.browser.newPage();

    // Disable timeout for slower devices
    this.page.setDefaultNavigationTimeout(0);
    this.page.setDefaultTimeout(0);

    this.page.goto(CLIENT_URL, {
      waitUntil: ['load', 'networkidle2'],
    });

    await this.page.waitForNavigation();
    await this.page.waitForTimeout(20000);
    const signedIn = this.page.url().includes(CLIENT_URL);

    if (!signedIn) {
      await this.page.waitForTimeout(1000);
      await this.browser.close();
      await this.login();
      this.initialize();
    } else {
      this.initialized = true;
      console.log('Outlook initialized.');
    }
  }

  getData() {
    if (this.initialized) {
      return this.getThreads();
    } else {
      return of({
        data: [{ message: 'Outlook feed not initialized', icon: ICON_PATH }],
      });
    }
  }

  login() {
    return new Promise(async (resolve) => {
      this.browser = await getBrowser({
        headless: false,
        userDataDir: './puppeteer-outlook-session',
      });
      this.page = await this.browser.newPage();
      // Disable timeout for slower devices
      this.page.setDefaultNavigationTimeout(0);
      this.page.setDefaultTimeout(0);

      console.log('Opening sign in page...');
      await this.page.goto(CLIENT_URL, {
        waitUntil: ['load', 'networkidle2'],
      });

      console.log('Please login to Outlook');
      await this.page.waitForFunction(
        `window.location.href.includes('https://outlook.office.com/mail')`,
      );
      await this.page.waitForTimeout(30000);

      await this.browser.close();

      console.log('Outlook signed in!');
      resolve(true);
    });
  }

  async getThreads(): Promise<any> {
    return new Promise(async (resolve, reject) => {
      await (async () => {
        try {
          const threadList = await this.page.waitForSelector(
            '[aria-label="Message list"]',
          );
          const threads = await threadList.$$('[role="option"]');
          if (threads.length) {
            const items = [];
            for (const thread of threads) {
              const unread = await thread.evaluate((el) =>
                el.getAttribute('aria-label'),
              );

              if (unread.indexOf('Unread') === 0) {
                const threadText = unread;
                const name = threadText.split('\n')[0];
                items.push({ message: name, icon: ICON_PATH });
              }
            }
            resolve({ data: items });
          }
        } catch (e) {
          reject(e);
        }
      })();
    });
  }
}
