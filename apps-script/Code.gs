/**
 * 최지인 Data Analyst – portfolio web app
 * Page source: Drive folder "portfolio-web" / index.html (replace that file to update the site).
 * Images: run setupImages() once (and again after adding portfolio-pf-*.zip files).
 */
const SRC_FOLDER = 'portfolio-web';
const PAGE_FILE = 'index.html';
const IMG_FOLDER = 'portfolio-web-pf';
const ZIP_PREFIX = 'portfolio-pf-';

function doGet() {
  const t = HtmlService.createTemplate(getPage_());
  t.imgJson = JSON.stringify(getImageMap_());
  return t.evaluate()
    .setTitle('최지인 Data Analyst')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getPage_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('page');
  if (hit) return hit;
  const folder = DriveApp.getFoldersByName(SRC_FOLDER).next();
  const files = folder.getFilesByName(PAGE_FILE);
  let best = null;
  while (files.hasNext()) {
    const f = files.next();
    if (f.isTrashed()) continue;
    if (!best || f.getLastUpdated() > best.getLastUpdated()) best = f;
  }
  const html = best.getBlob().getDataAsString('UTF-8');
  try { cache.put('page', html, 300); } catch (e) {}
  return html;
}

function getImageMap_() {
  const raw = PropertiesService.getScriptProperties().getProperty('IMG_MAP');
  return raw ? JSON.parse(raw) : {};
}

function setupImages() {
  const it = DriveApp.getFoldersByName(IMG_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(IMG_FOLDER);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const map = {};
  const existing = folder.getFiles();
  while (existing.hasNext()) { const f = existing.next(); if (!f.isTrashed()) map[f.getName()] = f.getId(); }
  const zips = DriveApp.searchFiles("title contains '" + ZIP_PREFIX + "' and trashed = false");
  while (zips.hasNext()) {
    const z = zips.next();
    if (!/\.zip$/i.test(z.getName())) continue;
    Utilities.unzip(z.getBlob().setContentType('application/zip')).forEach(function (b) {
      const name = b.getName().split('/').pop();
      if (!/\.(jpe?g|png)$/i.test(name) || map[name]) return;
      b.setContentType(/\.png$/i.test(name) ? 'image/png' : 'image/jpeg').setName(name);
      map[name] = folder.createFile(b).getId();
    });
  }
  Object.keys(map).forEach(function (n) {
    try { DriveApp.getFileById(map[n]).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  });
  PropertiesService.getScriptProperties().setProperty('IMG_MAP', JSON.stringify(map));
  CacheService.getScriptCache().remove('page');
  Logger.log('Images ready: ' + Object.keys(map).length);
}
