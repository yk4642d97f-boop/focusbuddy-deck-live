// 专注搭子 · WPS 演示插入器
// Puts the live 3D film into slide 1 as a WPS web control (WebShape), full bleed.
// Auto-runs once for a file whose name contains "生成器", then saves "<same folder>/专注搭子-放映稿.pptx".
var BASE = location.href.replace(/wps\/.*$/, '');
var DECK_URL = BASE;
var MARK = '生成器';

function app() {
  return wps.WppApplication ? wps.WppApplication() : wps.Application;
}

function insertDeck(pres) {
  var ps = pres.PageSetup;
  var W = ps.SlideWidth, H = ps.SlideHeight;
  var slide = pres.Slides.Count ? pres.Slides.Item(1) : pres.Slides.Add(1, 12);
  for (var i = slide.Shapes.Count; i >= 1; i--) {
    var s0 = slide.Shapes.Item(i);
    if (s0.Name === 'FocusBuddyDeck' || s0.Name === 'FocusBuddyPoster') s0.Delete();
  }
  // still frame underneath, shown while the page loads
  try { slide.Shapes.AddPicture(BASE + 'wps/poster.jpg', 0, -1, 0, 0, W, H).Name = 'FocusBuddyPoster'; } catch (e) {}
  var sh = slide.Shapes.AddWebShapeEx(DECK_URL);
  try { sh.LockAspectRatio = 0; } catch (e) {}
  sh.Left = 0;
  sh.Top = 0;
  sh.Width = W;
  sh.Height = H;
  try { sh.Name = 'FocusBuddyDeck'; } catch (e) {}
  return sh;
}

function note(pres, msg) {
  try {
    var slide = pres.Slides.Item(1);
    var tb = slide.Shapes.AddTextbox(1, 20, 20, 900, 200);
    tb.TextFrame.TextRange.Text = msg;
  } catch (e) {}
}

var OUT = '专注搭子-放映稿.pptx';
function saveOut(pres, suffix) {
  var name = OUT.replace('.pptx', (suffix || '') + '.pptx');
  var targets = ['/Users/yangcong/Desktop/' + name, (pres.Path || '') + '/' + name, '/Users/yangcong/Library/Containers/com.kingsoft.wpsoffice.mac/Data/Documents/' + name];
  for (var i = 0; i < targets.length; i++) {
    try { pres.SaveCopyAs(targets[i]); } catch (e) {}
  }
  try { pres.Save(); } catch (e) {}
}

var tries = 0, done = {};
function autoRun() {
  tries++;
  try {
    var a = app();
    var n = a.Presentations.Count;
    for (var i = 1; i <= n; i++) {
      var p = a.Presentations.Item(i);
      var name = p.Name || '';
      if (name.indexOf(MARK) >= 0 && !done[name]) {
        done[name] = 1;
        try {
          insertDeck(p);
          saveOut(p);
        } catch (e) {
          note(p, '插入失败: ' + (e && (e.message || e.description || e)));
          saveOut(p, '-错误');
        }
      }
    }
  } catch (e) {}
  if (tries < 90) setTimeout(autoRun, 1000);
}

function OnAddinLoad(ribbonUI) {
  if (typeof wps.ribbonUI !== 'object') wps.ribbonUI = ribbonUI;
  setTimeout(autoRun, 1500);
  return true;
}

function OnInsertDeck(control) {
  try {
    insertDeck(app().ActivePresentation);
  } catch (e) {
    alert('插入失败：' + (e && (e.message || e)));
  }
  return true;
}
