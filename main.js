const puppeteer = require("puppeteer");
require("dotenv").config();

const verbose = process.env.VERBOSE;
const log = (msg) => {
  if (verbose) console.log(`${new Date().toISOString()}: ${msg}\n`);
};
let history = [];

(async () => {
  async function delay(time) {
    log(`Waiting for navigation + ${time}ms`);
    await page.waitForNavigation();
    return new Promise(function (res) {
      setTimeout(res, time);
    });
  }

  const browser = await puppeteer.launch({
    args: ["--window-size=640,1600"],
    defaultViewport: null,
    headless: "new",
  });
  const page = await browser.newPage();
  await page.goto(process.env.SHAREPOINT);
  await delay(500);

  // login if not logged in
  if (await page.$("#loginHeader > div")) {
    log("Logging in");
    log("Entered email");
    await page.type("#i0116", process.env.EMAIL);
    await page.click("#idSIButton9");

    await delay(500);
    log("Entered password");
    await page.type("#i0118", process.env.PASSWORD);
    await page.click("#idSIButton9");

    // "Stay signed in?" prompt
    await delay(500);
    await page.click("#idSIButton9");
  }

  let name = "";
  await page.waitForSelector("#O365_MainLink_Me");
  name = await page.evaluate(
    (el) => el.textContent.slice(0, -2),
    await page.$("#O365_MainLink_Me")
  );
  log(`Logged in as ${name}`);

  const customIcon = (name) => {
    const avatars = [
      ["André Sousa", 47536659],
      ["João Alves", 59509896],
      ["Marco Porto", 50577988],
      ["Matthew Rodrigues", 38044816],
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

    const now = new Date();
    const date = new Date(now.getTime());

    if (str.includes("second"))
      date.setSeconds(date.getSeconds() - parseInt(str));
    else if (str.includes("minute"))
      date.setMinutes(date.getMinutes() - parseInt(str));
    else if (str.includes("hour"))
      date.setHours(date.getHours() - parseInt(str));

    // unix timestamp: date.getTime()
    return date.toISOString();
  };

  async function getUpdates() {
    await page.waitForSelector('[data-automationid="detailsPane"]');
    await page.click('[data-automationid="detailsPane"]');
    log("Opened details pane");

    await page.waitForSelector(".od-ItemActivityFeed");
    log("Details pane loaded");

    try {
      await page.waitForSelector('[aria-label="Today"]');
    } catch {
      log("No updates today yet");
      return;
    }
    const todayActivities = await page.$$(
      '[aria-label="Today"] > div > div > div'
    );

    for (const activity of todayActivities) {
      let deleted = false,
        renamed = 0;

      const items = await activity.$$(".ms-ActivityItem-activityContent > div");

      let text = await page.evaluate((el) => el.textContent, items[0]);
      if (text.includes("deleted")) deleted = true;
      if (text.includes("rename")) renamed = 1;

      let links = await items[0].$$("a");
      let idx = 0;
      for (const link of links) {
        const content = await link.$eval("span", (el) => el.textContent.trim());

        if (content.includes(".url")) continue; // ignore .url files
        
        await activity.scrollIntoView(); // this isn't doing jack

        await page.keyboard.down("Control");
        await link.click();
        await page.keyboard.up("Control");

        const pages = await browser.pages(),
          t = pages[pages.length - 1],
          url = t.url();
        await t.close();

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

        text = replaceAt(text, content, idx, `[${content}](${url})`);
        idx += replacement.length + url.length + 4;
      }
      text = text.replace(/  /g, " ");

      // name match
      const nMatch = /^(.*?)\s*(?:create|delete|edit|move|rename|share)/i.exec(
        text
      );
      let author = nMatch && nMatch[1];
      text = text.replace(author, "").trim();
      if (author == "You") author = name;

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
      if (date.includes("Yesterday")) return;
      const timestamp = parseTimeAgo(date);

      if (history.includes(`${author}${text}`)) return;

      log(`Sending embed: ${author} - ${text}`);
      sendDiscordEmbed(author, text, timestamp);
      history.push(`${author}${text}`);
    }
  }
  await getUpdates();

  setInterval(async () => {
    log("Reloading page");
    await page.reload({ waitUntil: ["networkidle0", "domcontentloaded"] });
    getUpdates();
  }, process.env.INTERVAL * 60000);
})();
