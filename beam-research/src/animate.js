// Adds click-by-click PowerPoint animations to slides pptxgenjs has written, which pptxgenjs cannot do itself.
//
// steps: { [slideNumber]: [click, ...] } where a click is [group, ...] and a group is
// { delay?: ms after the click, items: [shapeName | { name, effect: "fade" | "wipe" | "path", dur?, path? }] }.
// Shapes are found by the objectName given in pptxgenjs. The first group of a click starts on the click; later
// groups start automatically after their delay. "path" moves a shape along a motion path in slide-relative units.
const fs = require("fs");

function loadJSZip() {
  return require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
}

function timingXml(clicks, ids) {
  let n = 2; // cTn ids 1 and 2 are the root and main sequence
  const next = () => ++n;
  const grp = {}; // spid -> next build group id
  const builds = [];

  const effectXml = (item, nodeType) => {
    const it = typeof item === "string" ? { name: item } : item;
    const spid = ids[it.name];
    if (!spid) throw new Error(`animate: no shape named "${it.name}"`);
    const effect = it.effect || "fade";
    const dur = it.dur || 500;
    const g = (grp[spid] = (grp[spid] ?? -1) + 1);
    builds.push(`<p:bldP spid="${spid}" grpId="${g}" animBg="1"/>`);
    const tgt = `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;
    const id = next();
    if (effect === "path") {
      return `<p:par><p:cTn id="${id}" presetID="0" presetClass="path" presetSubtype="0" accel="50000" decel="50000" fill="hold" grpId="${g}" nodeType="${nodeType}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
        `<p:animMotion origin="layout" path="${it.path}" pathEditMode="relative" ptsTypes="AA"><p:cBhvr><p:cTn id="${next()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>` +
        `</p:childTnLst></p:cTn></p:par>`;
    }
    const [preset, sub, filter] = effect === "wipe" ? [22, 8, "wipe(left)"] : [10, 0, "fade"];
    return `<p:par><p:cTn id="${id}" presetID="${preset}" presetClass="entr" presetSubtype="${sub}" fill="hold" grpId="${g}" nodeType="${nodeType}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
      `<p:set><p:cBhvr><p:cTn id="${next()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
      `<p:animEffect transition="in" filter="${filter}"><p:cBhvr><p:cTn id="${next()}" dur="${dur}"/>${tgt}</p:cBhvr></p:animEffect>` +
      `</p:childTnLst></p:cTn></p:par>`;
  };

  const clickXml = clicks.map((groups) => {
    const outer = next();
    const inner = groups.map((group, gi) => {
      const parId = next();
      const items = group.items.map((item, ii) =>
        effectXml(item, ii > 0 ? "withEffect" : gi === 0 ? "clickEffect" : "afterEffect")).join("");
      return `<p:par><p:cTn id="${parId}" fill="hold"><p:stCondLst><p:cond delay="${gi === 0 ? 0 : group.delay || 0}"/></p:stCondLst><p:childTnLst>${items}</p:childTnLst></p:cTn></p:par>`;
    }).join("");
    return `<p:par><p:cTn id="${outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;
  }).join("");

  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${clickXml}</p:childTnLst></p:cTn>` +
    `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
    `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>` +
    `</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>${builds.join("")}</p:bldLst></p:timing>`;
}

async function addAnimations(file, steps) {
  const JSZip = loadJSZip();
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const [num, clicks] of Object.entries(steps)) {
    const part = `ppt/slides/slide${num}.xml`;
    let xml = await zip.file(part).async("string");
    const ids = {};
    for (const m of xml.matchAll(/<p:cNvPr id="(\d+)" name="([^"]*)"/g)) {
      if (ids[m[2]]) throw new Error(`animate: duplicate shape name "${m[2]}" on slide ${num}`);
      ids[m[2]] = m[1];
    }
    if (xml.includes("<p:timing>")) throw new Error(`animate: slide ${num} already has timing`);
    xml = xml.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + timingXml(clicks, ids));
    zip.file(part, xml);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  if (process.env.ANIM_DUMP) {
    // Shape names revealed by each click, for rendering preview frames of every build state.
    const names = (clicks) => clicks.map((groups) => groups.flatMap((g) => g.items.map((i) => (typeof i === "string" ? i : i.name))));
    fs.writeFileSync(process.env.ANIM_DUMP, JSON.stringify(Object.fromEntries(Object.entries(steps).map(([k, v]) => [k, names(v)]))));
  }
}

module.exports = { addAnimations };
