(function () {
  "use strict";
  var data = window.SummitHelpData || { groups: [] };
  var overrides = window.SummitHelpOverrides || {};
  var navigation = document.getElementById("navigation");
  var content = document.getElementById("content");
  var search = document.getElementById("search");

  function escapeHtml(value) {
    return String(value == null ? "" : value).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  function query() {
    var result = {}, parts = window.location.search.replace(/^\?/, "").split("&"), i, pair;
    for (i = 0; i < parts.length; i += 1) {
      if (parts[i]) {
        pair = parts[i].split("=");
        result[decodeURIComponent(pair[0])] = decodeURIComponent((pair.slice(1).join("=") || "").replace(/\+/g, " "));
      }
    }
    return result;
  }

  function findPage(gsid, csid, name) {
    var g, c, i, j;
    for (i = 0; i < data.groups.length; i += 1) {
      g = data.groups[i];
      for (j = 0; j < g.children.length; j += 1) {
        c = g.children[j];
        if ((String(g.id) === String(gsid) && String(c.id) === String(csid)) ||
            (!gsid && name && c.name.toLowerCase() === name.toLowerCase())) return { group: g, child: c };
      }
    }
    return null;
  }

  function allText(child) {
    var text = child.name + " " + (child.defaultWorksheet || "") + " " + (child.userGuide || ""), i, j, s, f;
    for (i = 0; i < child.sections.length; i += 1) {
      s = child.sections[i]; text += " " + s.name;
      for (j = 0; j < s.fields.length; j += 1) { f = s.fields[j]; text += " " + f.name + " " + (f.tip || ""); }
    }
    return text.toLowerCase();
  }

  function renderNavigation(filter) {
    var html = "", i, j, g, c, needle = (filter || "").toLowerCase();
    for (i = 0; i < data.groups.length; i += 1) {
      g = data.groups[i]; html += '<div class="nav-group">' + escapeHtml(g.name) + "</div>";
      for (j = 0; j < g.children.length; j += 1) {
        c = g.children[j];
        if (!needle || allText(c).indexOf(needle) >= 0) html += '<button class="nav-item" data-gsid="' + g.id + '" data-csid="' + c.id + '">' + escapeHtml(c.name) + "</button>";
      }
    }
    navigation.innerHTML = html;
  }

  function renderHome() {
    content.innerHTML = '<article class="page"><h1>Summit Help</h1><p class="eyebrow">Workbook-backed guidance for Abovo Summit</p><div class="intro"><p>Select an interface from the contents, or search for a field, section, worksheet, or topic.</p><p>This first help library is assembled from the authoritative workbook User Guide, Excel cell comments, and Summit\'s editable Structure.xml presentation definitions.</p></div><div class="notice"><strong>Editing this help</strong><p>Generated workbook facts live in <code>Help\\data\\interfaces.js</code>. Put durable client wording in <code>Help\\data\\overrides.js</code>; that file is deliberately not replaced by the generator. The <em>Open help folder</em> button above opens these files.</p></div><h2>Working safely</h2><ul><li>Use Summit controls for ordinary inputs so changes are validated, calculated, and recorded in history.</li><li>Use <strong>Add lines</strong> for structural changes; it coordinates every workbook range required by that rule.</li><li>Review Check Sheet warnings before closing or saving a model.</li><li>The XLSB remains the authoritative model and must continue to round-trip through Microsoft Excel and VBA.</li></ul></article>';
  }

  function renderPage(found, requestedSection) {
    var g = found.group, c = found.child, key = g.id + ":" + c.id, o = overrides[key] || {}, i, j, s, f;
    var html = '<article class="page"><h1>' + escapeHtml(c.name) + '</h1><p class="eyebrow">' + escapeHtml(g.name) + ' · Summit interface</p>';
    if (o.summary || c.userGuide) html += '<div class="intro">' + (o.summary ? o.summary : '<p>' + escapeHtml(c.userGuide) + '</p>') + '</div>';
    html += '<div class="meta"><span class="pill">Default worksheet: ' + escapeHtml(c.defaultWorksheet || "Not specified") + '</span><span class="pill">' + c.sections.length + ' section' + (c.sections.length === 1 ? '' : 's') + '</span></div>';
    if (o.additionalHtml) html += o.additionalHtml;
    html += '<h2>Sections and fields</h2>';
    for (i = 0; i < c.sections.length; i += 1) {
      s = c.sections[i]; html += '<section class="section" id="section-' + i + '"><h3>' + escapeHtml(s.name) + '</h3>';
      if (s.dataSources.length) html += '<p class="muted">Data: ' + escapeHtml(s.dataSources.join(', ')) + '</p>';
      for (j = 0; j < s.fields.length; j += 1) {
        f = s.fields[j]; html += '<div class="field"><div class="field-name">' + escapeHtml(f.name) + '</div>' + (f.tip ? '<div>' + escapeHtml(f.tip) + '</div>' : '') + '</div>';
      }
      if (!s.fields.length) html += '<p class="empty">No field-level guidance is currently available for this section.</p>';
      html += '</section>';
    }
    if (c.comments.length) {
      html += '<h2>Workbook notes</h2>';
      for (i = 0; i < c.comments.length; i += 1) html += '<div class="source-note"><strong>' + escapeHtml(c.comments[i].sheet + ' ' + c.comments[i].address) + '</strong><div>' + escapeHtml(c.comments[i].text) + '</div></div>';
    }
    html += '<p class="muted">Source: Structure.xml and read-only metadata from TestFileClean.xlsb. Client-authored overrides take precedence.</p></article>';
    content.innerHTML = html;
    if (requestedSection) {
      for (i = 0; i < c.sections.length; i += 1) {
        if (c.sections[i].name.toLowerCase() === requestedSection.toLowerCase()) { document.getElementById('section-' + i).scrollIntoView(); break; }
      }
    }
  }

  function navigate(gsid, csid, name, section) { var found = findPage(gsid, csid, name); if (found) renderPage(found, section); else renderHome(); }
  navigation.onclick = function (e) { var target = e.target || e.srcElement; if (target && target.getAttribute("data-csid") !== null) navigate(target.getAttribute("data-gsid"), target.getAttribute("data-csid"), "", ""); };
  document.getElementById("home-link").onclick = renderHome;
  search.onkeyup = function () { renderNavigation(search.value); };
  renderNavigation("");
  var q = query(); navigate(q.gsid, q.csid, q.name, q.section);
}());
