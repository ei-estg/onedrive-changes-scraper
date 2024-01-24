const puppeteer = require("puppeteer");
require("dotenv").config();
const { sub, formatISO } = require("date-fns");

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

  let browser = await puppeteer.launch({ headless: "new" });
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

  customIcon = (name) => {
    const avatars = [
      ["André Sousa", 47536659],
      ["João Alves", 59509896],
      ["Marco Porto", 50577988],
      ["Matthew Rodrigues", 38044816],
    ];
    if (avatars.some((avatar) => avatar[0] == name))
      return `https://avatars.githubusercontent.com/u/${
        avatars.find((avatar) => avatar[0] == name)[1]
      }?s=100&v=4`;
    else
      return "https://cdn.discordapp.com/attachments/1045218965439389698/1199670717378203658/u.png?ex=65c3636b&is=65b0ee6b&hm=03cc6d239ea38d15bf0d9aec6ca0b91ca055a3f50a36e998a0fee3b2177c57df&";
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

  async function getUpdates() {
    await page.waitForSelector('[data-automationid="detailsPane"]');
    await page.click('[data-automationid="detailsPane"]');
    log("Opened details pane");

    await page.waitForSelector(".od-ItemActivityFeed");
    log("Details pane loaded");

    await page.waitForSelector('[aria-label="Today"]');
    const todayActivities = await page.$$(
      '[aria-label="Today"] > div > div > div'
    );

    for (const activity of todayActivities) {
      const items = await activity.$$(".ms-ActivityItem-activityContent > div");

      let text = await page.evaluate((el) => el.textContent, items[0]);

      // name match
      const nMatch = /^(.*?)\s*(?:create|delete|edit|rename|share)/i.exec(text);
      let author = nMatch && nMatch[1];
      text = text.replace(author, "").trim();
      if (author == "You") author = name;

      // translate some text elems
      const dict = [
        ["created", "Criou"],
        ["deleted", "Apagou"],
        ["edited", "Editou"],
        ["from", "de"],
        ["in", "na pasta"],
        ["renamed", "Alterou o nome de"],
        ["shared", "Partilhou"],
        ["to", "para"],
      ];
      dict.forEach(([original, word]) => {
        const regex = new RegExp(`\\b${original}\\b`, "gi");
        text = text.replace(regex, word);
      });

      let date = await page.evaluate((el) => el.textContent, items[1]);
      if (date.includes("About") || date.includes("Yesterday")) return;
      const timestamp = formatISO(
        sub(new Date(), {
          [date.split(" ")[1].toLowerCase()]: parseInt(date),
        })
      );

      if (history.includes(`${author}${text}`)) return;

      log(`Sending embed: ${author} - ${text}`);
      sendDiscordEmbed(author, text, timestamp);
      history.push(`${author}${text}`);
    }
  }
  await getUpdates();

  setInterval(async () => {
    await page.reload({ waitUntil: ["networkidle0", "domcontentloaded"] });
    getUpdates();
  }, process.env.INTERVAL * 60000);
})();
