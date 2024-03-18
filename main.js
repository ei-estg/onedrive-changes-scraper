require("dotenv").config();
const os = require("os");
const fs = require("fs");
const puppeteer = require("puppeteer");
const cron = require("node-cron");

const log = (msg, forced) => {
  if (process.env.VERBOSE !== "false" || forced)
    console.log(`${new Date().toISOString()}: ${msg}\n`);
};
const err = (msg, forced) => {
  log(`\x1b[31m${msg}\x1b[0m`, forced);
};

// Delete puppeteer profiles from tmp directory to free up space
// github.com/stefanzweifel/sidecar-browsershot/pull/54/files
if (os.platform() === "linux") {
  fs.readdirSync("/tmp").forEach((file) => {
    if (file.startsWith("puppeteer_dev_chrome_profile"))
      fs.rm(`/tmp/${file}`, { recursive: true });
  });
  log("Deleted puppeteer cache /tmp", 1);
}

const waitForSelector = async (page, selector) => {
  try {
    const startTime = Date.now();

    const element = await page.waitForSelector(selector);
    log(`Selected ${selector} element within ${Date.now() - startTime} ms`);

    return element;
  } catch (error) {
    err(`Waiting for selector ${selector} failed`);
  }
};

(async () => {
  const delay = async (time) => {
    log(`Waiting for navigation + ${time}ms`);
    try {
      await page.waitForNavigation();
    } catch {
      err("No navigation to wait for?");
    }
    return new Promise((res) => {
      setTimeout(res, time);
    });
  };

  const browser = await puppeteer.launch({
    args: ["--force-device-scale-factor=0.4", "--window-size=900,2400"],
    defaultViewport: null,
    executablePath: process.env.BROWSER || undefined,
    headless: process.env.DEBUG == "true" ? false : "new",
  });

  // set download path
  const client = await browser.target().createCDPSession();
  const bakPath = process.env.BACKUP_DIR;
  await client.send("Browser.setDownloadBehavior", {
    behavior: "allowAndName",
    downloadPath: bakPath,
    eventsEnabled: true,
  });
  client.on("Browser.downloadProgress", async (evt) => {
    if (evt.state === "completed")
      fs.renameSync(
        `${bakPath}/${evt.guid}`,
        `${bakPath}/${new Date().getTime()}.zip`
      );
  });

  const page = await browser.newPage();
  await page.goto(process.env.SHAREPOINT);
  await delay(500);

  // login if not logged in
  if (await page.$("#loginHeader > div")) {
    log("Logging in");
    await page.type("#i0116", process.env.EMAIL);
    log("Entered email");
    await page.click("#idSIButton9");
    await delay(500);

    await page.type("#i0118", process.env.PASSWORD);
    log("Entered password");
    await page.click("#idSIButton9");
    await delay(500);

    // "Stay signed in?" prompt
    await page.click("#idSIButton9");
  }

  await waitForSelector(page, "#O365_MainLink_Me");
  let name = await page.evaluate(
    (el) => el.textContent.slice(0, -2),
    await page.$("#O365_MainLink_Me")
  );
  log(`Logged in as ${name}`);

  const replaceAt = (str, target, start, replacement) => {
    const idx = str.indexOf(target, start);
    return idx !== -1
      ? str.slice(0, idx) + replacement + str.slice(idx + target.length)
      : str;
  };

  const parseTimeAgo = (str) => {
    str = str.replace("A few seconds", "30 seconds");
    str = str.replace("About a minute", "1 minute");
    str = str.replace("About an hour", "1 hour");

    const date = new Date();

    if (str.includes("second"))
      date.setSeconds(date.getSeconds() - parseInt(str));
    else if (str.includes("minute"))
      date.setMinutes(date.getMinutes() - parseInt(str));
    else if (str.includes("hour"))
      date.setHours(date.getHours() - parseInt(str));

    // unix timestamp: date.getTime()
    return date.toISOString();
  };

  const customIcon = (name) => {
    const avatars = [
      ["André Sousa", 47536659],
      ["João Alves", 59509896],
      ["Marco Porto", 50577988],
      ["Matthew Rodrigues", 38044816],
      ["Pedro Cunha", 72658683],
      ["Rodrigo Sá", 22347167],
    ];
    return avatars.some((avatar) => avatar[0] == name)
      ? `https://avatars.githubusercontent.com/u/${
          avatars.find((avatar) => avatar[0] == name)[1]
        }?s=100&v=4`
      : "https://cdn.discordapp.com/attachments/1045218965439389698/1199670717378203658/u.png?ex=65c3636b&is=65b0ee6b&hm=03cc6d239ea38d15bf0d9aec6ca0b91ca055a3f50a36e998a0fee3b2177c57df&";
  };

  const sendDiscordEmbed = (author, msg, timestamp) => {
    fetch(process.env.DISCORD_WEBHOOK, {
      headers: {
        "content-type": "application/json",
      },
      body: `{"embeds":[{"description":"${msg}","color":2664682,"author":{"name":"${author}","url":"https://ipvcpt-my.sharepoint.com/personal/amatossousa_ipvc_pt/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Famatossousa%5Fipvc%5Fpt%2FDocuments%2FEI%5FESTG%5FIPVC&ga=1","icon_url":"${customIcon(
        author
      )}"},"timestamp":"${timestamp}"}],"username":"OneDrive EI","avatar_url":"https://cdn.discordapp.com/attachments/1045218965439389698/1199729748234997895/o.png?ex=65c39a65&is=65b12565&hm=34986c78561219d35fb0165badf1442ae378273a47f2eb77f72ffa8a4d6f6bc9&"}`,
      method: "POST",
    });
  };

  let history = [];
  async function getUpdates() {
    try {
      await waitForSelector(page, '[data-automationid="detailsPane"]');
      await page.click('[data-automationid="detailsPane"]');
      log("Opened details pane");

      await waitForSelector(page, ".od-ItemActivityFeed");
      log("Details pane loaded");

      try {
        await waitForSelector(page, '[aria-label="Today"]');
      } catch {
        log("No updates today yet");
        return;
      }
      const todayActivities = await page.$$(
        '[aria-label="Today"] > div > div > div'
      );

      for (const activity of todayActivities.reverse()) {
        const items = await activity.$$(
          ".ms-ActivityItem-activityContent > div"
        );

        let text = await page.evaluate((el) => el.textContent, items[0]);

        // name match
        const nMatch =
          /^(.*?)\s*(?:create|delete|edit|move|rename|share)/i.exec(text);
        let author = nMatch && nMatch[1];
        text = text.replace(author, "").trim();
        if (author == "You") author = name;

        if (history.includes(`${author}${text}`)) continue;
        history.push(`${author}${text}`);

        let deleted = false,
          renamed = 0;

        if (text.includes("deleted")) deleted = true;
        if (text.includes("rename")) renamed = 1;

        let links = await items[0].$$("a");
        let idx = 0;
        for (const link of links) {
          const content = await link.$eval("span", (el) => el.title);

          let url = "";

          // ignore .url and .zip files
          if (!/\.(url|zip)/.test(content)) {
            await page.keyboard.down("Control");
            await link.click();
            await page.keyboard.up("Control");

            const pages = await browser.pages(),
              t = pages[pages.length - 1];
            url = t.url();
            await t.close();
          }

          const replacement =
            renamed !== 2
              ? url.includes("file") || content.includes(".")
                ? `the file **${content}**`
                : `the folder **${content}**`
              : `**${content}**`;
          text = replaceAt(text, content, idx, replacement);

          if (deleted) {
            text = text.replace(/deleted (.*) from/, `deleted **$1** from`);
            text = text.replace(content, `[${content}](${url})`);
            break;
          }

          if (renamed == 1) {
            renamed++;
            continue;
          }

          text = replaceAt(
            text,
            content,
            idx,
            url ? `[${content}](${url})` : `**${content}**`
          );
          idx += replacement.length + url.length + 4;
        }
        text = text.replace(/  /g, " ");

        // uppercase first letter from text
        text = text.charAt(0).toUpperCase() + text.slice(1);

        // translate text elems to Portuguese
        const dict = [
          ["Created", "Criou"],
          ["Deleted", "Eliminou"],
          ["Edited", "Editou"],
          ["Moved", "Moveu"],

          ["Renamed the file", "Alterou o nome do ficheiro"],
          ["Renamed the folder", "Alterou o nome da pasta"],

          ["Shared", "Partilhou"],

          ["in the folder", "na pasta"],
          ["from the folder", "da pasta"],
          ["the folder", "a pasta"],

          ["the file", "o ficheiro"],

          ["to", "para"],
        ];
        dict.forEach(([original, word]) => {
          const regex = new RegExp(`\\b${original}\\b`, "gi");
          text = text.replace(regex, word);
        });

        let date = await page.evaluate((el) => el.textContent, items[1]);
        if (date.includes("Yesterday")) continue;
        const timestamp = parseTimeAgo(date);

        log(`Sending embed: ${author} - ${text}`, 1);
        sendDiscordEmbed(author, text, timestamp);
      }
    } catch (e) {
      err("Error (Most likely internet connection issue): " + e, 1);
    }
  }
  await getUpdates();

  cron.schedule(process.env.INTERVAL, async () => {
    log("Reloading page", 1);
    try {
      await page.reload({ waitUntil: ["networkidle0", "domcontentloaded"] });
    } catch (e) {
      err("Error (Most likely internet connection issue): " + e, 1);
    }
    getUpdates();
  });

  if (process.env.BACKUP !== "false") {
    if (process.env.BACKUP_QNTY < 1)
      return err(
        "'BACKUP_QNTY' var can't be 0. The script will skip the backup procedure.",
        1
      );
    else if (process.env.BACKUP_QNTY == 1)
      err(
        `'BACKUP_QNTY' var is set to 1! The script WILL overwrite the ONLY backup if not stopped beforehand. Proceed cautiously!!`,
        1
      );
    cron.schedule(process.env.BACKUP_FREQ, async () => {
      try {
        await waitForSelector(page, '[data-automationid="downloadCommand"]');

        const files = fs.readdirSync(bakPath);
        if (files.length >= process.env.BACKUP_QNTY) {
          const oldest = files.sort((a, b) => a - b)[0];
          fs.unlinkSync(`${bakPath}/${oldest}`);
          log(`Replaced oldest backup: ${oldest}`, 1);
        }

        page.click('[data-automationid="downloadCommand"]');
        log("Downloading backup", 1);
      } catch (e) {
        err("Couldn't download backup: " + e, 1);
      }
    });
  }
})();
