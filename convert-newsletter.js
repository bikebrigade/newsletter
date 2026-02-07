const { execSync } = require('child_process');
require('dotenv').config();
const { google } = require('googleapis');
const glob = require('glob');
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');
const cheerio = require('cheerio');
const axios = require('axios');
const sharp = require('sharp');
const USE_CACHE = true;

const DRIVE_ID = process.env.DRIVE_ID;
const MAILCHIMP_API_KEY = process.env.MAILCHIMP_KEY;
const MAILCHIMP_HEADER = 'Basic ' + Buffer.from('anystring:' + MAILCHIMP_API_KEY).toString('base64');
const MAILCHIMP_SERVER = process.env.MAILCHIMP_SERVER || 'us1';
const BRIGADE_PREVIEW = process.env.BRIGADE_PREVIEW;
const DOWNLOAD_DIR = process.env.DOWNLOAD_DIR;
const TEMP_DIR = process.env.TEMP_DIR;
const OUTPUT_HTML_FILE = 'test.html';
const BRIGADE_NEWSLETTER_IMG_DIR = process.env.IMAGE_ARCHIVE;

const nextDate = getNextNewsletterDate();

function getNextNewsletterDate() {
  const today = new Date();
  const todayDayOfWeek = today.getDay();
  const daysUntilSunday = (7 - todayDayOfWeek) % 7 || 7;
  today.setDate(today.getDate() + daysUntilSunday);
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  const formattedDate = `${year}-${month}-${day}`;
  return formattedDate;
}

async function getLatestNewsletterSketchpad() {
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/drive.readonly'],
  });
  const drive = google.drive({ version: 'v3', auth });
  const listRes = await drive.files.list({
    q: "name contains 'newsletter sketchpad' and mimeType = 'application/vnd.google-apps.document' and trashed = false",
    spaces: 'drive',
    corpora: 'drive',
    driveId: DRIVE_ID,
    includeItemsFromAllDrives: true,
    supportsAllDrives: true,
    orderBy: 'modifiedTime desc',
    pageSize: 1,
    fields: 'files(id, name, parents)',
  });
  if (listRes.data.files.length === 0) throw new Error('No newsletter file found.');
  return listRes.data.files[0];
}

async function copyPreviousNewsletter(latestFile) {
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/documents'],
  });
  const drive = google.drive({ version: 'v3', auth });
  const docs = google.docs({ version: 'v1', auth });
  const copyRes = await drive.files.copy({
    fileId: latestFile.id,
    requestBody: {
      name: `${nextDate} newsletter sketchpad`,
      parents: latestFile.parents,
    },
    supportsAllDrives: true,
  });
  const newFileId = copyRes.data.id;
  await drive.permissions.create({
    fileId: newFileId,
    requestBody: { role: 'writer', type: 'anyone' },
    supportsAllDrives: true,
  });
  return docs.documents.get({ documentId: newFileId });
}

async function updateEditLink(doc) {
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/documents'],
  });
  const docs = google.docs({ version: 'v1', auth });
  const editLink = `https://docs.google.com/document/d/${doc.data.documentId}/edit`;
  const requests = [];
  const searchText = "Draft/notes in this Google Doc";
  doc.data.body.content.forEach(element => {
    if (element.paragraph) {
      element.paragraph.elements.forEach(el => {
        if (el.textRun && el.textRun.content.includes(searchText)) {
          const startIndex = el.startIndex;
          const endIndex = startIndex + searchText.length;

          requests.push({
            updateTextStyle: {
              range: { startIndex, endIndex },
              textStyle: {
                link: { url: editLink }
              },
              fields: 'link'
            }
          });
        }
      });
    }
  });
  if (requests.length > 0) {
    await docs.documents.batchUpdate({
      documentId: doc.data.documentId,
      requestBody: { requests }
    });
  }
}

async function processAndFixImages(tempDir) {
  const imagesDir = path.join(tempDir, 'images');
  const htmlPath = getHTMLFilename(tempDir);

  if (!fs.existsSync(imagesDir) || !htmlPath) return;

  const htmlContent = fs.readFileSync(htmlPath, 'utf8');
  const $ = cheerio.load(htmlContent, { decodeEntities: false });
  const files = fs.readdirSync(imagesDir);

  for (const file of files) {
    const ext = path.extname(file).toLowerCase();
    if (!['.png', '.jpeg', '.jpg', '.webp', '.gif'].includes(ext)) continue;

    const inputPath = path.join(imagesDir, file);
    const $img = $(`img[src="images/${file}"]`);
    const alt = $img.attr('alt') || '';

    const isEmoji = /^:.+:$/.test(alt);
    const fileNameNoExt = path.parse(file).name;
    const outputExt = isEmoji ? '.png' : '.jpg';
    const outputFileName = `${fileNameNoExt}${outputExt}`;
    const outputPath = path.join(imagesDir, outputFileName);

    try {
      let pipeline = sharp(inputPath).resize({ width: 800, withoutEnlargement: true });

      if (isEmoji) {
        pipeline = pipeline.png({ palette: true });
      } else {
        pipeline = pipeline.jpeg({ quality: 85 });
      }

      const buffer = await pipeline.toBuffer();
      fs.writeFileSync(outputPath, buffer);
      if (file !== outputFileName) {
        fs.unlinkSync(inputPath);
      }
    } catch (err) {
      console.error(`Failed processing ${file}:`, err.message);
    }
  }
}

async function maybeCreateNewsletterPad() {
  const latestFile = await getLatestNewsletterSketchpad();
  const latestDate = latestFile.name.substring(0, 10);
  if (latestDate < nextDate) {
    console.log(`Creating new sketchpad for ${nextDate}...`);
    try {
      const doc = await copyPreviousNewsletter(latestFile);
      await updateEditLink(doc);
      return doc.documentId;
    } catch (err) {
      console.log(err);
    }
  } else {
    console.log(`There is a current newsletter for ${nextDate}`);
    console.log('https://docs.google.com/document/d/' + latestFile.id + '/edit');
    return false;
  }
}

function getHTMLFilename(dir) {
  const htmlFiles = glob.sync(`${dir}/*.html`).sort((a, b) => {
    return fs.statSync(b).mtime - fs.statSync(a).mtime;
  });
  return htmlFiles[0];
}

async function recentMailchimpFiles() {
  const response = await axios.get(
    `https://${MAILCHIMP_SERVER}.api.mailchimp.com/3.0/file-manager/files?count=100&sort_field=added_date&sort_dir=DESC`,
    { headers: { 'Authorization': MAILCHIMP_HEADER } }
  );
  return response.data;
}

async function uploadImagesToMailchimp(sectionMap) {
  const files = await recentMailchimpFiles();
  const fileMap = {};
  const base = function (s) {
    return s.replace(/^.+?-news-|\.(jpg|png)$/g, '');
  };
  const maybeUploadImage = async function (filename) {
    if (filename.match(/^data/)) return filename;
    if (fileMap[base(filename)]) {
      return fileMap[base(filename)];
    }
    const fileData = fs.readFileSync(filename).toString('base64');
    const response = await axios.post(
      `https://${MAILCHIMP_SERVER}.api.mailchimp.com/3.0/file-manager/files`,
      { name: path.basename(filename), file_data: fileData },
      { headers: { Authorization: `apikey ${MAILCHIMP_API_KEY}` } }
    );
    return response.data.full_size_url;
  };
  for (let f of files.files) {
    fileMap[base(f.name)] = f.full_size_url;
  }
  for (let section of sectionMap.results) {
    for (let img of section.images) {
      img.url = await maybeUploadImage(img.filename);
    }
  }
  return sectionMap;
}

async function parseSections(tempDir) {
  const htmlPath = getHTMLFilename(tempDir);
  const htmlContent = fs.readFileSync(htmlPath, 'utf8');
  const $ = cheerio.load(htmlContent);

  const sectionMap = {};
  let pendingImages = [];
  let lastHeadingText = 'intro';
  $('*').each((i, el) => {
    const $el = $(el);
    if ($el.is('img')) {
      const src = $el.attr('src');
      if (pendingImages.length > 0) {
        pendingImages[pendingImages.length - 1].extra = true;
      }
      if (src && src.startsWith('images/')) {
        pendingImages.push({ src: src, alt: $el.attr('alt') });
      }
    }
    else if ($el.is('h1') && pendingImages.length > 0) {
      pendingImages[pendingImages.length - 1].extra = true;
      processPendingImages(pendingImages, lastHeadingText, tempDir, sectionMap);
      pendingImages = [];
    } else if ($el.is('h2, h3')) {
      const currentHeading = $el.text()
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '-')
        .substring(0, 30);
      if (pendingImages.length > 0) {
        processPendingImages(pendingImages, currentHeading, tempDir, sectionMap);
        pendingImages = [];
      }
      lastHeadingText = currentHeading;
    }
  });
  if (pendingImages.length > 0) {
    processPendingImages(pendingImages, lastHeadingText, tempDir, sectionMap);
  }
  return sectionMap;
}

async function processPendingImages(imageSrcs, sectionName, tempDir, sectionMap) {
  for (let i = 0; i < imageSrcs.length; i++) {
    const localSrc = imageSrcs[i];
    const localFilePath = path.join(tempDir, localSrc.src);
    if (!fs.existsSync(localFilePath)) continue;
    const extension = path.extname(localSrc.src);
    newFilename = `${nextDate}-news-${sectionName}${(localSrc.extra) ? '-extra' : ''}${extension}`;
    sectionMap[sectionName] ||= {};
    mapEntry = sectionMap[sectionName];
    if (localSrc.extra) {
      mapEntry.extra = { local: localFilePath, name: newFilename, alt: localSrc.alt || sectionName };
    } else {
      sectionMap[sectionName] = { ...mapEntry, local: localFilePath, name: newFilename, alt: localSrc.alt || sectionName };
    }
    return sectionMap;
  }
}

async function transformNewsletter() {
  try {
    if (!USE_CACHE) {
      const zipPath = await downloadLatestNewsletter();
      const zip = new AdmZip(zipPath);
      zip.extractAllTo(TEMP_DIR, true);
      await processAndFixImages(TEMP_DIR);
    }
  } catch (err) {
    if (err.response) console.error('API Details:', err.response.data);
  }
}

const ZIP_PATH = 'newsletter.zip';

async function downloadNewsletterByFile(fileId, tempDir) {
  const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/drive.readonly'],
  });
  const drive = google.drive({ version: 'v3', auth });

  if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });
  const imagesDir = path.join(tempDir, 'images');
  if (!fs.existsSync(imagesDir)) fs.mkdirSync(imagesDir, { recursive: true });

  console.log('📄 Exporting HTML content...');
  const htmlRes = await drive.files.export({
    fileId: fileId,
    mimeType: 'text/html'
  });
  const htmlContent = htmlRes.data;
  const htmlPath = path.join(tempDir, 'newsletter.html');
  fs.writeFileSync(htmlPath, htmlContent);
  const docs = google.docs({ version: 'v1', auth });
  const doc = await docs.documents.get({ documentId: fileId });
  const inlineObjects = doc.data.inlineObjects || {};
  const objectIds = Object.keys(inlineObjects);
  const imageMap = {}; // Maps internal doc ID to local filename

  for (let i = 0; i < objectIds.length; i++) {
    const objId = objectIds[i];
    const internalSource = inlineObjects[objId].inlineObjectProperties.embeddedObject.imageProperties.contentUri;
    const response = await axios({
      method: 'get',
      url: internalSource,
      responseType: 'stream'
    });

    const localFileName = `image${i + 1}.png`;
    const localPath = path.join(imagesDir, localFileName);
    const writer = fs.createWriteStream(localPath);

    response.data.pipe(writer);

    await new Promise((resolve, reject) => {
      writer.on('finish', resolve);
      writer.on('error', reject);
    });
    imageMap[objId] = `images/${localFileName}`;
  }

  return { htmlPath, imageMap };
}

async function downloadLatestNewsletter() {
  const file = await getLatestNewsletterSketchpad();
  console.log(`Downloading ${file.name} (${file.id})`);
  const res = await downloadNewsletterByFile(file.id, TEMP_DIR);
}

function getNextWeekday(date, dayOfWeek, strictlyAfter = true) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  let diff = (dayOfWeek - d.getDay() + 7) % 7;
  if (strictlyAfter && diff === 0) diff = 7;
  d.setDate(d.getDate() + diff);
  return d;
}

function formatDate(date, format) {
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const pad = (n) => String(n).padStart(2, '0');

  switch (format) {
    case 'ISO': return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
    case 'MONTH_DAY': return `${months[date.getMonth()]} ${date.getDate()}`;
    case 'DAY_ONLY': return `${date.getDate()}`;
    case 'MONTH_ONLY': return `${months[date.getMonth()]}`;
    default: return "";
  }
}

function getRangeStr(start, end) {
  const startStr = formatDate(start, 'MONTH_DAY');
  const endStr = (start.getMonth() === end.getMonth())
    ? formatDate(end, 'DAY_ONLY')
    : formatDate(end, 'MONTH_DAY');
  return `${startStr}-${endStr}`;
}

function makeBrigadeSignupBlock(inputDate = new Date()) {
  const baseDate = getNextWeekday(inputDate, 0, false);
  const currentWeek = getNextWeekday(baseDate, 1, true);
  const currentWeekEnd = getNextWeekday(baseDate, 0, true);
  const nextWeek = getNextWeekday(currentWeek, 1, true);
  const nextWeekEnd = getNextWeekday(currentWeekEnd, 0, true);
  const curISO = formatDate(currentWeek, 'ISO');
  const nextISO = formatDate(nextWeek, 'ISO');
  const curRange = getRangeStr(currentWeek, currentWeekEnd).toUpperCase();
  const nextRange = getRangeStr(nextWeek, nextWeekEnd);
  return `<table class="sign-up" style="background-color: #223f4d; text-align: center; margin: auto; margin-top: 24px; margin-bottom: 12px;"><tbody><tr><td><a href="https://dispatch.bikebrigade.ca/campaigns/signup?current_week=${curISO}" target="_blank" class="sign-up mceButtonLink" style="background-color:#223f4d;border-radius:0;border:2px solid #223f4d;color:#ffffff;display:block;font-family:'Helvetica Neue', Helvetica, Arial, Verdana, sans-serif;font-size:16px;font-weight:normal;font-style:normal;padding:16px 28px;text-decoration:none;text-align:center;direction:ltr;letter-spacing:0px" rel="noreferrer">SIGN UP NOW TO DELIVER ${curRange}</a></td></tr></table>
<p style="text-align: center; font-family: 'Helvetica Neue', Helvetica, Arial, Verdana"><a href="https://dispatch.bikebrigade.ca/campaigns/signup?current_week=${nextISO}" style="color: #476584; margin-top: 12px; margin-bottom: 12px;" target="_blank">You can also sign up early to deliver ${nextRange}</a></p>`;
}

function slugify(s) {
  return s.trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function saveImages($, dir, filePrefix = "", transformFn = null) {
  let lastImage = null; // { ext, data (base64) }
  let lastImageFilename = null;
  let lastImageAlt = "";
  let lastImageNode = null;
  let results = [];
  $('img, h2').each((i, el) => {
    const node = $(el);
    const tag = el.name;
    if (tag === 'img') {
      let data = node.attr('src') || "";
      let alt = node.attr('alt') || "";

      if (!data.startsWith('data:')) {
        const ext = path.extname(data);
        const jpgPath = data.replace(new RegExp(`\\${ext}$`), '.jpg');
        if (fs.existsSync(jpgPath)) {
          node.attr('src', jpgPath);
          data = jpgPath;
        }
      }

      if (lastImage || lastImageFilename) {
        const currentSection = results[0];
        const sectionName = currentSection ? (transformFn ? transformFn(currentSection.section) : currentSection.section) : "unknown";
        let extraFilename = "";

        if (lastImage) {
          extraFilename = path.join(dir, `${filePrefix}${sectionName}-extra.${lastImage.ext}`);
          fs.writeFileSync(extraFilename, Buffer.from(lastImage.data, 'base64'));
        } else if (lastImageFilename) {
          extraFilename = path.join(dir, `${filePrefix}${sectionName}-extra.jpg`);
          try {
            execSync(`convert "${lastImageFilename}" "${extraFilename}"`);
          } catch (e) {
            fs.copyFileSync(lastImageFilename, extraFilename);
          }
        }

        if (currentSection) {
          currentSection.images.push({
            type: 'extra',
            filename: extraFilename,
            alt: lastImageAlt
          });
        }

        if (lastImageNode) $(lastImageNode).remove();
        lastImage = null;
        lastImageFilename = null;
        lastImageNode = null;
      }

      if (alt.match(/^:.+?:/)) {
        if (results.length > 0) {
          results[0].images.push({
            type: 'emoji',
            filename: data,
            alt: alt
          });
        }
      } else if (data.startsWith('images/')) {
        lastImage = null;
        lastImageFilename = data;
        lastImageAlt = alt;
        lastImageNode = el;
      } else if (data.startsWith('data:image/')) {
        const match = data.match(/^data:image\/([^;]+);base64,(.*)$/);
        if (match) {
          lastImage = { ext: match[1], data: match[2] };
          lastImageFilename = null;
          lastImageAlt = alt;
          lastImageNode = el;
        }
      }
    }

    else if (tag === 'h2') {
      const text = node.text().trim();
      if (text !== "") {
        const slug = transformFn ? transformFn(text) : text;
        let finalFilename = null;

        if (lastImage) {
          finalFilename = path.join(dir, `${filePrefix}${slug}.${lastImage.ext}`);
          fs.writeFileSync(finalFilename, Buffer.from(lastImage.data, 'base64'));
        } else if (lastImageFilename) {
          const ext = path.extname(lastImageFilename);
          finalFilename = path.join(dir, `${filePrefix}${slug}${ext}`);
          fs.copyFileSync(lastImageFilename, finalFilename);
        }

        results.unshift({
          section: text,
          images: finalFilename ? [{ type: 'main', filename: finalFilename, alt: lastImageAlt}] : []
        });


        if (lastImageNode) $(lastImageNode).remove();
        lastImage = null;
        lastImageFilename = null;
        lastImageNode = null;
      }
    }
  });

  return results.reverse();
}

function myTransformHtmlRemoveItalics($) {
  $('i').remove();
  return $;
}

/**
 * Extracts CSS rules from <style> tags into a Map
 */
function myHtmlExtractCssRules($, elem) {
  const cssRules = new Map();
  const cssContent = elem.find('style').text();
  const regex = /([^{]+)\{([^}]+)\}/g;
  let match;

  while ((match = regex.exec(cssContent)) !== null) {
    const selector = match[1].trim();
    const properties = match[2].trim();
    cssRules.set(selector, properties);
  }
  return cssRules;
}

function myBrigadeConvertSpanStyle(node, rules, $) {
  const className = $(node).attr('class');
  const text = $(node).text().trim();

  if (className && text !== "") {
    const classes = className.split(/\s+/);
    const styles = classes
      .map(c => rules.get(`.${c}`))
      .filter(prop => prop && (prop.includes('font-weight:700') || prop.includes('font-style:italic')))
      .join(';');

    if (styles) {
      $(node).attr('style', styles);
      $(node).removeAttr('class');
      return true;
    }
  }
  return false;
}

function myBrigadeSimplifyHtml($, elem) {
  const rules = myHtmlExtractCssRules($, elem);
  const tagsToClean = ['li', 'b', 'ul', 'span', 'p', 'a', 'h2', 'div'];
  tagsToClean.forEach(tag => {
    elem.find(tag).each((i, node) => {
      const converted = myBrigadeConvertSpanStyle(node, rules, $);
      if (!converted) {
        if ($(node).attr('style') != 'font-weight:700') { $(node).removeAttr('style'); }
        $(node).removeAttr('class');
        $(node).removeAttr('id');
      }
    });
  });
  $('sup').remove();
  $('p').each((i, node) => {
    const hasImage = $(node).find('img').length > 0;
    if ($(node).text().trim() === "" && !hasImage) {
      $(node).remove();
    }
  });
  $('a').each((i, node) => {
    const href = $(node).attr('href');
    if (href && href.includes("https://www.google.com/url")) {
      try {
        const parsedUrl = new URL(href);
        const realUrl = parsedUrl.searchParams.get('q');
        if (realUrl) {
          $(node).attr('href', realUrl);
        }
      } catch (e) {
      }
    }
  });

  return $;
}


/**
 * Convert [ Button Label ] links into styled HTML tables for email.
 * This looks for links inside paragraphs, removes the paragraph,
 * and inserts the button table structure before it.
 */
function myBrigadeFormatButtons($, item) {
  item.find('p').each((i, el) => {
    const $node = $(el);
    const text = $node.text().trim();
    const match = text.match(/^\[\s*(.+?)\s*\]$/);
    if (match) {
      const label = match[1];
      const href = $node.find('a').attr('href') || "";
      const buttonHtml = `<table style="margin: auto"><tbody><tr><td style="padding: 12px 0 12px 0"><div style="margin-top: 12px"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" class="mceButtonContainer" style="padding-top: 24px; margin: auto; margin-top: 12px; text-align: center"><tbody><tr class="mceStandardButton"><td style="background-color:#000000;border-radius:0;margin-top:12px;text-align:center" valign="top" class="mceButton"><a href="${href}" target="_blank" class="mceButtonLink" style="background-color:#000000;border-radius:0;border:2px solid #000000;color:#ffffff;display:block;font-family:'Helvetica Neue', Helvetica, Arial, Verdana, sans-serif;font-size:16px;font-weight:normal;font-style:normal;padding:16px 28px;text-decoration:none;text-align:center;direction:ltr;letter-spacing:0px" rel="noreferrer">${label}</a></td></tr></tbody></table></div></td></tr></tbody></table>`;
      if ($node.length > 0) {
        $node.before(buttonHtml);
        $node.remove();
      }
    }
  });
  return $;
}

const MY_BRIGADE_COMMUNITY_TEXT_STYLE = "color: #ffffff";
const MY_BRIGADE_COMMUNITY_LINK_STYLE = "color: #aed9ef";

/**
 * Change the colors of links and text in a DOM tree recursively.
 * - Tables are skipped.
 * - Links (except .mceButtonLink) get a specific style.
 * - Text nodes are wrapped in <span> with a specific style.
 */
function myBrigadeRecolorRecursively($, node) {
  const tagName = node.name;
  if (tagName === 'table') {
    return;
  }
  if (tagName === 'a') {
    const $node = $(node);
    if (!$node.hasClass('mceButtonLink')) {
      $node.attr('style', MY_BRIGADE_COMMUNITY_LINK_STYLE);
    }
    return;
  }
  if (node.childNodes) {
    const children = [...node.childNodes];
    children.forEach((child) => {
      if (child.type === 'text') {
        const textContent = $(child).text();
        if (textContent.trim() !== "") {
          $(child).replaceWith(`<span style="${MY_BRIGADE_COMMUNITY_TEXT_STYLE}">${textContent}</span>`);
        }
      } else if (child.type === 'tag') {
        myBrigadeRecolorRecursively($, child);
      }
    });
  }
}

function processRecolor($, elem) {
  elem.children().each((i, el) => {
    myBrigadeRecolorRecursively($, el);
  });
  return $;
}

function myBrigadeJustHeadings($) {
  const $ul = $('<ul></ul>');
  $('h2').each((i, el) => {
    const text = $(el).text().trim();
    if (text !== "") {
      $ul.append($('<li></li>').text(text));
    }
  });
  return $ul.prop('outerHTML');
}

function getNextSundayStr() {
  const d = new Date();
  d.setDate(d.getDate() + (7 - d.getDay()) % 7);
  return d.toISOString().substring(0, 10);
}

function saveNewsletterImages($) {
  const datePrefix = `${getNextSundayStr()}-news-`;

  const results = saveImages(
    $,
    BRIGADE_NEWSLETTER_IMG_DIR,
    datePrefix,
    slugify
  );

  return {
    results,
    updatedHtml: $.html()
  };
}


/**
 * Divides a list of nodes into sections based on a specific tag (e.g., 'h2').
 * Returns an array of objects: { section: string, children: node[] }
 */
function myHtmlGroupByTag($, nodes, tag) {
  const results = [];
  let currentSectionName = null;
  let currentChildren = [];

  nodes.each((i, node) => {
    const $node = $(node);
    if (node.name === tag && $node.text().trim() !== "") {
      if (currentSectionName !== null) {
        results.push({ section: currentSectionName, children: [...currentChildren] });
      }
      currentSectionName = $node.text().trim();
      currentChildren = [];
    } else if (currentSectionName !== null) {
      currentChildren.push(node);
    }
  });

  if (currentSectionName !== null) {
    results.push({ section: currentSectionName, children: currentChildren });
  }

  return results;
}

/**
 * Generates Table of Contents <li> items from section groups.
 * Looks for date patterns (e.g., "Sun Dec 1") in the first child of a section.
 */
function myBrigadeTocItems($, nodes) {
  const groups = myHtmlGroupByTag($, $(nodes), 'h2');
  const dayNames = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  const dateRegex = new RegExp(`^(${dayNames.join('|')})\\s+([A-Za-z]+\\s+[0-9]+)`);

  return groups.map(group => {
    const sectionName = group.section;
    const firstChild = group.children[0];
    let label = sectionName;

    if (firstChild) {
      const firstChildText = $(firstChild).text().trim();
      const match = firstChildText.match(dateRegex);
      if (match) {
        label = `${match[2]}: ${sectionName}`;
      }
    }

    const escapedLabel = label
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");

    return `<li>${escapedLabel}</li>`;
  });
}

/**
 * Generates the Table of Contents HTML.
 */
function myBrigadeToc($, sections) {
  const bikeBrigadeItems = sections["Bike Brigade"]
        ? myBrigadeTocItems($, sections["Bike Brigade"])
    : [];
  const communityItems = sections["In our community"]
        ? myBrigadeTocItems($, sections["In our community"])
    : [];
  const allItems = [...bikeBrigadeItems, ...communityItems];
  const ulHtml = `<ul>${allItems.join('')}</ul>`;
  return ulHtml.replace(/<li>/g, '\n<li>');
}

function myBrigadeBlock(text, options = {}) {
  const bg = options.bg || "#223f4d";
  const style = options.style || "padding-left:24px;padding-right:24px;padding-top:12px;padding-bottom:12px";

  return `<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse" role="presentation"><tbody><tr><td style= "padding-top:0;padding-bottom:0;padding-right:0;padding-left:0;border:0;border-radius:0" valign="top"><table width="100%" style= "border:0;background-color:${bg};border-radius:0"><tbody><tr><td style="${style}" class="mceTextBlockContainer"><div data-block-id="738" class="mceText" style= "width:100%">${text}</div></td></tr></tbody></table></td></tr></tbody></table>`;
}

/**
 * Formats a specific section (list of nodes) into the newsletter layout
 */
function myBrigadeFormatSection($, sectionNodes, imageMap, recolor = false) {
  const groups = myHtmlGroupByTag($, $(sectionNodes), 'h2');

  return groups.map(group => {
    const heading = group.section;
    const children = group.children;
    const sectionImages = imageMap.results.find((o) => o.section == heading)?.images || [];
    const mainImage = sectionImages.find(img => img.type === 'main');
    const extraImage = sectionImages.find(img => img.type === 'extra');
    const $item = $('<div></div>').append($(`<h2>${heading}</h2>`)).append($(children));

    let callToAction = null;
    $item.find('p').each((i, el) => {
      if (/^\[ *.+ *\]/.test($(el).text().trim())) {
        callToAction = $(el).find('a').attr('href');
      }
    });
    myBrigadeSimplifyHtml($, $item);
    myBrigadeFormatButtons($, $item);
    if (recolor) {
      processRecolor($, $item);
    }
    let contentHtml = $item.html();

    const extraImageHtml = extraImage
      ? `<tr><td colspan="2" style="padding-top: 12px"><img src="${extraImage.url}" alt="${extraImage.alt || ""}" style="width: 100%; max-width: 100%"></td></tr>`
      : "";

    if (mainImage) {
      const imageAlt = mainImage.alt || heading;
      const imageElement = callToAction
        ? `<a href="${callToAction}" tabindex="-1" style="display: block;"><span style="background-color: transparent"><img src="${mainImage.url}" alt="${imageAlt}" style="padding-top: 12px; display:block;max-width:100%;height:auto;border-radius:0" width="306" height="auto" class="imageDropZone mceImage"></span></a>`
        : `<img src="${mainImage.url}" alt="${imageAlt}" style="display:block; padding-top: 12px; width:100%; max-width:100%;height:auto;border-radius:0" width="306" height="auto" class="imageDropZone mceImage">`;

      return `<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" style="margin-top: 12px; margin-bottom: 12px;"><tbody><tr class="mceRow"><td colspan="1" rowspan="1" style="background-position:center;background-repeat:no-repeat;background-size:cover" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0"><tbody><tr><td colspan="12" rowspan="1" valign="top" width="100%" class="mceColumn"><table width="100%" border="0" cellspacing="0" cellpadding="0"><tbody><tr><td colspan="1" rowspan="1" style="border:0;border-radius:0" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" align="center"><tbody><tr class="mceRow"><td colspan="1" rowspan="1" style="background-position:center;background-repeat:no-repeat;background-size:cover;padding-top:0px;padding-bottom:0px" valign="top"><table style="table-layout:fixed" width="100%" border="0" cellspacing="24" cellpadding="0"><tbody><tr><td colspan="6" rowspan="1" style="padding-top:0;padding-bottom:0" valign="top" width="50%" class="mceColumn"><table width="100%" border="0" cellspacing="0" cellpadding="0"><tbody><tr><td colspan="1" rowspan="1" style="background-color:transparent;border:0;border-radius:0" valign="top" class="mceImageBlockContainer"><table style="border-collapse:separate;margin:0;vertical-align:top;max-width:100%;width:100%;height:auto" width="100%" border="0" cellspacing="0" cellpadding="0" align="center"><tbody><tr><td colspan="1" rowspan="1" style="border:0;border-radius:0;margin:0" valign="top">${imageElement}</td></tr></tbody></table></td></tr></tbody></table></td><td colspan="6" rowspan="1" style="padding-top:0;padding-bottom:0" valign="top" width="50%" class="mceColumn"><table width="100%" border="0" cellspacing="0" cellpadding="0"><tbody><tr><td colspan="1" rowspan="1" style="padding:12px" valign="top" class="mceGutterContainer"><table style="border-collapse:separate" width="100%" border="0" cellspacing="0" cellpadding="0"><tbody><tr><td colspan="1" rowspan="1" style="padding-top:0;padding-bottom:0;padding-right:0;padding-left:0;border:0;border-radius:0" valign="top"><table style="border:0;background-color:transparent;border-radius:0;border-collapse:separate" width="100%"><tbody><tr><td colspan="1" rowspan="1" class="mceTextBlockContainer">${contentHtml}</td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr>${extraImageHtml}</tbody></table>`;
    } else {
      return `<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" style="margin-top: 12px; margin-bottom: 12px;"><tbody><tr><td colspan="2" style="padding-top: 12px">${contentHtml}</td></tr>${extraImageHtml}</tbody></table>`;
    }
  }).join("");
}

/**
 * Updates <img> sources based on filename mapping
 */
function myBrigadeUpdateImages($, images) {
  const flatImages = Object.values(images).flat();
  $('img').each((i, el) => {
    const src = $(el).attr('src');
    if (src) {
      const base = path.basename(src, path.extname(src));
      const match = flatImages.find(img => img.filename && path.basename(img.filename, path.extname(img.filename)) === base);
      if (match && match.url) {
        $(el).attr('src', match.url);
      }
    }
  });
}

/**
 * Main Process
 */
async function myBrigadeProcessNewsletter(htmlFile, date = new Date(getNextNewsletterDate())) {
  const rawHtml = fs.readFileSync(htmlFile, 'utf8');
  const $ = cheerio.load(rawHtml, { decodeEntities: true });
  let imageMetadata = saveNewsletterImages($);
  imageMetadata = await uploadImagesToMailchimp(imageMetadata);
  myBrigadeSimplifyHtml($, $.root());
  myBrigadeUpdateImages($, imageMetadata);
  const bodyChildren = $('body').children();
  const sectionsRaw = myHtmlGroupByTag($, bodyChildren, 'h1');
  const sections = {};
  sectionsRaw.forEach(s => sections[s.section] = s.children);
  const dateStr = date.toLocaleDateString('en-US', { timeZone: 'UTC', month: 'long', day: 'numeric', year: 'numeric' });

  let html = `<table class="newsletter" margin=0 cellpadding=0 cellspacing=0 style="border-collapse:collapse"><tbody><tr><td>`;
  html += myBrigadeBlock(`<table style="margin: auto"><tbody><tr><td style="text-align: center; color: #f3f3f3"><div style="text-align: center; color: #f3f3f3">${dateStr}</div></td></tr></tbody></table>`, {
    bg: "#16232a",
    style: "padding: 0px 24px 12px 24px"
  });

  html += `<base href=""><style>.tpl-content { padding: 0 !important } table { border-collapse: collapse !important } table.newsletter { border-collapse: collapse} .mceStandardButton a, table.sign-up a { text-decoration: none }</style><table><tbody><tr><td style="padding: 12px 24px 12px 24px"><p>Hi Bike Brigaders! Here's what's happening this week, with quick signup links. In this e-mail:</p>`;
  html += myBrigadeToc($, sections);

  html += makeBrigadeSignupBlock(date);

  if (sections["Bike Brigade"]) {
    html += myBrigadeFormatSection($, sections["Bike Brigade"], imageMetadata);
  }

  html += `</td></tr></tbody></table>`;

  if (sections["In our community"]) {
    html += `<table style="background-color:#223f4d;"><tbody><tr><td style="padding-left: 24px; padding-right: 24px">`;
    html += myBrigadeBlock(`<h1 style="text-align: center;"><span style= "color:#ffffff;">In our community</span></h1>`);
    html += myBrigadeFormatSection($, sections["In our community"], imageMetadata, true);
    html += `</td></tr></tbody></table>`;
  }

  if (sections["Other updates"]) {
    html += `<table><tbody><tr><td style="padding: 12px 24px 12px 24px"><h2>Other updates</h2><ul>${$(sections["Other updates"]).html()}</ul></td></tr></tbody></table>`;
  }

  html += `</td></tr></tbody></table>`;

  return html.replace(/<p><span><\/span><\/p>/g, "");
}

async function mailchimpRequest(endpoint, method = 'GET', data = null) {
  const response = await axios({
    method,
    url: `https://${MAILCHIMP_SERVER}.api.mailchimp.com/3.0${endpoint}`,
    data,
    headers: {
      Authorization: `apikey ${MAILCHIMP_API_KEY}`,
      'Content-Type': 'application/json'
    }
  });
  return response.data;
}

function getLatestFile(dir, ext) {
  if (!fs.existsSync(dir)) {
    return null;
  }

  const files = fs.readdirSync(dir);

  const sortedFiles = files
    .filter(file => file.toLowerCase().endsWith(ext.toLowerCase()))
    .map(file => {
      const filePath = path.join(dir, file);
      const stats = fs.statSync(filePath);
      return {
        path: filePath,
        mtime: stats.mtime.getTime()
      };
    })
    .sort((a, b) => b.mtime - a.mtime);

  return sortedFiles.length > 0 ? sortedFiles[0].path : null;
}

async function myBrigadeSchedule() {
  let campaign = await nextCampaign();
  let campaignId = campaign.id;
  const date = new Date(nextDate.replace(/-/g, '/') + ' 11:00:00').toISOString();
  if (campaign.send_time && new Date(campaign.send_time).toISOString() == date) {
    console.log(`Already scheduled for ${date}`);
  } else {
    try {
      if (campaign.send_time) {
        await mailchimpRequest(`/campaigns/${campaignId}/actions/unschedule`, 'POST');
      }
      await mailchimpRequest(`/campaigns/${campaignId}/actions/schedule`, 'POST', {
        schedule_time: date
      });
      console.log(`Scheduled campaign ${campaignId} for ${scheduleTime}`);
    } catch (err) {
      const errorMsg = err.response ? JSON.stringify(err.response.data) : err.message;
      console.error(`Failed to schedule campaign: ${errorMsg}`);
    }
  }
}

async function nextCampaign() {
  const campaignList = await mailchimpRequest('/campaigns?count=10&sort_field=create_time&sort_dir=DESC');
  return campaignList.campaigns.find(c => c.settings.title === nextDate);

}

async function createCampaign() {
  const listName = "Bike Brigade";
  const lists = await mailchimpRequest('/lists');
  const list = lists.lists.find(l => l.name === listName);
  if (!list) throw new Error(`List "${listName}" not found`);
  return await mailchimpRequest('/campaigns', 'POST', {
    type: "regular",
    recipients: { list_id: list.id },
    settings: {
      title: nextDate,
      subject_line: "Bike Brigade: Weekly update",
      from_name: "Bike Brigade",
      reply_to: "info@bikebrigade.ca",
      tracking: { opens: true, html_clicks: true }
    }
  });
}

async function processNewsletterHTML(campaign) {
    await processAndFixImages(TEMP_DIR);
    const output = await myBrigadeProcessNewsletter(getHTMLFilename(TEMP_DIR), new Date(nextDate));
    const finalContent = await updateCampaignHTML(campaign, output);
    fs.writeFileSync(OUTPUT_HTML_FILE, finalContent.html);
    execSync(`scp "${OUTPUT_HTML_FILE}" "${BRIGADE_PREVIEW}"`, { stdio: 'inherit' });
}

async function myBrigadeCreateOrUpdateCampaign(useLocal = false) {
  try {
    let campaign = await nextCampaign();
    if (!campaign) {
      console.log("Creating new campaign...");
      campaign = await createCampaign();
    }
    if (useLocal) {
      if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
      const latestZip = getLatestFile(DOWNLOAD_DIR, ".zip");
      console.log(latestZip);
      execSync(`unzip -o "${latestZip}" -d "${TEMP_DIR}"`);
    } else {
      await downloadLatestNewsletter();
    }
    await processNewsletterHTML(campaign);
    await myBrigadeSchedule();
  } catch (err) {
    console.error("Error updating campaign:", err.response ? err.response.data : err.message);
    console.error("Stack trace:", err.stack);
  }
}

async function updateCampaignHTML(campaign, output) {
  const templateName = "Bike Brigade weekly update";
  const templates = await mailchimpRequest('/templates');
  const template = templates.templates.find(t => t.name === templateName);
  if (!template) throw new Error(`Template "${templateName}" not found`);
  await mailchimpRequest(`/campaigns/${campaign.id}/content`, 'PUT', {
    template: {
      id: template.id,
      sections: {
        "main_content_area": output
      }
    }
  });
  return mailchimpRequest(`/campaigns/${campaign.id}/content`);
}

async function main() {
  if (process.argv[2] == 'sketch' || process.argv[2] == 'copy') {
    await maybeCreateNewsletterPad();
  } else if (process.argv[2] == 'download') {
    await maybeCreateNewsletterPad();
    await downloadLatestNewsletter();
    await processAndFixImages(TEMP_DIR);
    console.log('Downloaded');
  } else if (process.argv[2] == 'resize') {
    await processAndFixImages(TEMP_DIR);
    console.log('Resized');
  } else if (process.argv[2] == 'update') {
    await myBrigadeCreateOrUpdateCampaign();
  } else if (process.argv[2] == 'from-zip') {
    await myBrigadeCreateOrUpdateCampaign(true);
  } else if (process.argv[2] == 'format') {
    await processNewsletterHTML(await nextCampaign());
  } else if (process.argv[2] == 'schedule') {
    await myBrigadeSchedule();
  } else {
    await myBrigadeCreateOrUpdateCampaign();
  }
}
main();
