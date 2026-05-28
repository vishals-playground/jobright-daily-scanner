// ============================================================
// JOBRIGHT DAILY SCANNER - Google Apps Script Version
// No local setup needed. Runs entirely in Google cloud.
// ------------------------------------------------------------
// SETUP STEPS:
//  1. Paste this entire file into your Apps Script editor
//  2. Run setOpenAiKey() once and enter your OpenAI API key
//  3. Run createScanSheet() once to create the output sheet
//  4. Deploy as Web App (Execute as: Me, Who has access: Only myself)
//  5. (Optional) Add a daily time-driven trigger on runScan()
// ============================================================

// ---------- PROFILE ----------
var PROFILE = {
  name: "Vishal S",
  target_roles: [
    "Software Engineer", "Senior Software Engineer", "Lead Engineer",
    "Solutions Architect", "Engineering Manager", "AI Engineer",
    "AI Architect", "Staff Engineer"
  ],
  years_experience: 10,
  management_years: 3,
  core_skills: [
    "Python", "C#", ".NET", "SQL", "REST API",
    "Microservices", "Distributed Systems", "System Design"
  ],
  ai_ml_skills: [
    "LangGraph", "LangChain", "RAG", "LLM Orchestration",
    "Agentic AI", "Multi-agent frameworks", "Prompt Engineering",
    "Vector Databases", "LLM API Integration", "OpenAI", "Anthropic",
    "NLP", "Retrieval-Augmented Generation", "AI Pipelines"
  ],
  cloud_infra: [
    "AWS", "EKS", "Kubernetes", "Docker", "CI/CD",
    "Containerization", "API Gateway", "Cloud Architecture",
    "Observability", "OpenTelemetry", "Distributed Tracing"
  ],
  databases: [
    "SQL Server", "PostgreSQL", "NoSQL", "Vector DB", "Pinecone", "Redis"
  ],
  patterns: [
    "CQRS", "Event-Driven Architecture", "Mediator Pattern",
    "Prototype Pattern", "SOLID", "Clean Architecture",
    "Microservices Architecture"
  ],
  soft_skills: [
    "Offshore Team Management", "Technical Leadership",
    "System Design", "Cross-functional Collaboration",
    "Architecture Decision Records"
  ],
  hard_disqualifiers: [
    "TypeScript", "Node.js", "Java", "Go", "Golang",
    "Ruby", "PHP", "Scala", "Django"
  ],
  domain_gaps: [
    "Healthcare", "HIPAA", "FinTech", "Crypto", "Web3",
    "Insurance", "FDA", "Blockchain"
  ],
  preferred_salary_min: 130000,
  work_authorization: "US Citizen or GC",
  location: "Remote preferred"
};

// ---------- CONFIG ----------

function setOpenAiKey() {
  var key = Browser.inputBox("Enter your OpenAI API Key:", Browser.Buttons.OK_CANCEL);
  if (key && key !== "cancel") {
    PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", key);
    Browser.msgBox("API key saved successfully.");
  }
}

function getOpenRouterKey() {
  var key = PropertiesService.getScriptProperties().getProperty('OPENROUTER_API_KEY');
  return key || '';
}

function getOpenAiKey() {
  var key = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key) throw new Error("OpenAI API key not set. Run setOpenAiKey() first.");
  return key;
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.create("Jobright Scanner Results");
  }
  var sheet = ss.getSheetByName("Scan Results");
  if (!sheet) {
    sheet = ss.insertSheet("Scan Results");
    sheet.appendRow([
      "Date", "Title", "Company", "Fit Score", "Verdict",
      "Matched Skills", "Gaps", "Gap Severity", "Rejection Reasons",
      "Tailored Summary", "Key Bullets", "Recommendation", "URL"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 13).setFontWeight("bold");
  }
  return sheet;
}

function createScanSheet() {
  getOrCreateSheet();
    Logger.log("Sheet 'Scan Results' is ready!");
}

// ---------- GMAIL ----------

function fetchJobrightEmails(daysBack) {
  daysBack = daysBack || 1;
  var query = "from:noreply@jobright.ai newer_than:" + daysBack + "d";
  var threads = GmailApp.search(query);
  var bodies = [];
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      bodies.push(msg.getBody()); // HTML body
    });
  });
  return bodies;
}

// ---------- PARSING ----------

function extractJobLinks(htmlBody) {
  var links = [];
  var seen = {};
  var regex = /https:\/\/jobright\.ai\/jobs\/info\/[a-zA-Z0-9]+/g;
  var matches = htmlBody.match(regex);
  if (matches) {
    matches.forEach(function(url) {
      // Strip query params
      var clean = url.split("?")[0];
      if (!seen[clean]) {
        seen[clean] = true;
        links.push(clean);
      }
    });
  }
  return links;
}

// ---------- JOB FETCHING ----------

function fetchJobDetails(url) {
  try {
    var options = {
      headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0" },
      muteHttpExceptions: true,
      followRedirects: true
    };
    var response = UrlFetchApp.fetch(url, options);
    var html = response.getContentText();

    // Extract title from <h1>
    var titleMatch = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
    var title = titleMatch ? titleMatch[1].replace(/<[^>]+>/g, "").trim() : "Unknown Role";

    // Extract text from p, li, h2, h3, span tags
    var textBlocks = [];
    var tagRegex = /<(p|li|h2|h3|span)[^>]*>([\s\S]*?)<\/\1>/gi;
    var match;
    while ((match = tagRegex.exec(html)) !== null) {
      var text = match[2].replace(/<[^>]+>/g, "").replace(/&amp;/g, "&")
                         .replace(/&nbsp;/g, " ").replace(/&#\d+;/g, "").trim();
      if (text.length > 20) textBlocks.push(text);
    }

    return {
      url: url,
      title: title,
      text: textBlocks.slice(0, 300).join("\n")
    };
  } catch(e) {
    return { url: url, title: "Fetch Error", text: e.message };
  }
}

// ---------- AI ANALYSIS ----------

function callLLM(prompt) {
  var orKey = getOpenRouterKey();
  if (orKey) {
    try {
      var orPayload = {
        model: 'deepseek/deepseek-v4-flash',
        messages: [{ role: 'user', content: prompt }],
        temperature: 0.2,
        
      };
      var orOptions = {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Authorization': 'Bearer ' + orKey,
          'HTTP-Referer': 'https://script.google.com',
          'X-Title': 'Jobright Daily Scanner'
        },
        payload: JSON.stringify(orPayload),
        muteHttpExceptions: true
      };
      var orResp = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', orOptions);
      var orResult = JSON.parse(orResp.getContentText());
      if (orResult.choices && orResult.choices[0]) {
                Logger.log('Used DeepSeek V4 Flash via OpenRouter successfully');
        return JSON.parse(orResult.choices[0].message.content);
      }
      Logger.log('OpenRouter returned no choices, falling back to OpenAI...');
    } catch(e) {
      Logger.log('OpenRouter failed: ' + e.message + ' — falling back to OpenAI...');
    }
  } else {
    Logger.log('No OpenRouter key set, using OpenAI directly...');
  }
  // Fallback: OpenAI direct
  
  var apiKey = getOpenAiKey();
  var payload = {
    model: 'gpt-4o',
    messages: [{ role: 'user', content: prompt }],
    temperature: 0.2,
    
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
  var result = JSON.parse(response.getContentText());
  Logger.log('Used OpenAI fallback successfully');
  return JSON.parse(result.choices[0].message.content);
}

function analyzeFit(jobDetails) {
  var profileSummary = JSON.stringify({
    years_experience: PROFILE.years_experience,
    management_years: PROFILE.management_years,
    core_skills: PROFILE.core_skills,
    ai_ml_skills: PROFILE.ai_ml_skills,
    cloud_infra: PROFILE.cloud_infra,
    databases: PROFILE.databases,
    patterns: PROFILE.patterns,
    soft_skills: PROFILE.soft_skills,
    domain_gaps: PROFILE.domain_gaps,
    hard_disqualifiers: PROFILE.hard_disqualifiers
  }, null, 2);

  var prompt = "You are a senior technical recruiter and resume expert.\n\n" +
    "CANDIDATE PROFILE:\n" + profileSummary + "\n\n" +
    "JOB POSTING - " + jobDetails.title + ":\n" +
    jobDetails.text.substring(0, 3000) + "\n\n" +
    "Analyze the fit. Return a JSON object with these exact keys:\n" +
    "{\n" +
    '  "title": "job title",\n' +
    '  "company": "company name",\n' +
    '  "fit_score": 0-100 integer,\n' +
    '  "verdict": "STRONG_FIT" or "STRETCH" or "NOT_FIT",\n' +
    '  "matched_skills": ["list"],\n' +
    '  "gaps": ["list"],\n' +
    '  "gap_severity": "HARD" or "SOFT" or "NONE",\n' +
    '  "rejection_reasons": ["list if not fit"],\n' +
    '  "tailored_summary": "2-3 sentence resume summary tailored to this job using only real candidate skills",\n' +
    '  "key_bullets": ["3-5 resume bullet points emphasizing relevant experience"],\n' +
    '  "recommendation": "1-2 sentence advice"\n' +
    "}\n" +
    "Rules: fit_score 75+ = STRONG_FIT, 55-74 = STRETCH, below 55 = NOT_FIT. Never invent skills.";

    var analysis = callLLM(prompt);
  analysis.url = jobDetails.url;
  return analysis;
}

// ---------- MAIN SCAN ----------

function runScan() {
  Logger.log("Fetching Jobright emails from last 24 hours...");
  var emails = fetchJobrightEmails(1);
  Logger.log("Found " + emails.length + " Jobright email(s)");

  var allLinks = [];
  var seen = {};
  emails.forEach(function(body) {
    var links = extractJobLinks(body);
    links.forEach(function(link) {
      if (!seen[link]) {
        seen[link] = true;
        allLinks.push(link);
      }
    });
  });
  Logger.log("Found " + allLinks.length + " unique job postings");

  if (allLinks.length === 0) {
    Logger.log("No job links found. Check that Jobright alert emails arrived in the last 24h.");
    return;
  }

  var sheet = getOrCreateSheet();
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var results = [];

  allLinks.forEach(function(url, i) {
    Logger.log("Analyzing job " + (i + 1) + "/" + allLinks.length + ": " + url);
    try {
      var job = fetchJobDetails(url);
      Utilities.sleep(1000);
      var analysis = analyzeFit(job);
      results.push(analysis);

      // Write row to sheet immediately
      sheet.appendRow([
        dateStr,
        analysis.title || "",
        analysis.company || "",
        analysis.fit_score || 0,
        analysis.verdict || "",
        (analysis.matched_skills || []).join(", "),
        (analysis.gaps || []).join(", "),
        analysis.gap_severity || "",
        (analysis.rejection_reasons || []).join(", "),
        analysis.tailored_summary || "",
        (analysis.key_bullets || []).join(" | "),
        analysis.recommendation || "",
        url
      ]);

      Utilities.sleep(1000);
    } catch(e) {
      Logger.log("Error processing " + url + ": " + e.message);
    }
  });

  // Color-code rows by verdict
  colorCodeSheet(sheet);

  Logger.log("DONE! " + results.length + " jobs analyzed. Check 'Scan Results' sheet.");
  return results;
}

function colorCodeSheet(sheet) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var verdict = data[i][4];
    var range = sheet.getRange(i + 1, 1, 1, 13);
    if (verdict === "STRONG_FIT") {
      range.setBackground("#d9ead3"); // green
    } else if (verdict === "STRETCH") {
      range.setBackground("#fff2cc"); // yellow
    } else if (verdict === "NOT_FIT") {
      range.setBackground("#f4cccc"); // red
    }
  }
}

// ---------- WEB APP (optional dashboard) ----------

function doGet() {
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1).reverse(); // newest first

  var html = '<!DOCTYPE html><html><head><title>Jobright Scanner</title>' +
    '<style>' +
    'body{font-family:Arial,sans-serif;background:#11111b;color:#cdd6f4;padding:30px;max-width:1000px;margin:auto}' +
    'h1{color:#cba6f7}h2{color:#89b4fa;border-bottom:1px solid #313244;padding-bottom:8px}' +
    '.card{background:#1e1e2e;border-radius:10px;padding:18px;margin:12px 0}' +
    '.score{font-size:2em;font-weight:bold;float:right}' +
    '.STRONG_FIT{border-left:5px solid #a6e3a1}.STRETCH{border-left:5px solid #f9e2af}.NOT_FIT{border-left:5px solid #f38ba8}' +
    'a{color:#89b4fa}.tag{display:inline-block;background:#313244;border-radius:4px;padding:2px 8px;margin:2px;font-size:0.85em}' +
    '.btn{background:#a6e3a1;color:#11111b;border:none;padding:14px 36px;font-size:1.1em;border-radius:10px;cursor:pointer;font-weight:bold;margin-bottom:20px}' +
    '</style></head><body>' +
    '<h1>Jobright Daily Scanner</h1>' +
    '<form method="POST"><button class="btn" type="submit">Run Scan Now</button></form>';

  if (rows.length === 0) {
    html += '<p>No results yet.Run a scan first!</p>';
  } else {
    // Group by verdict
    var strong = rows.filter(function(r){ return r[4] === 'STRONG_FIT'; });
    var stretch = rows.filter(function(r){ return r[4] === 'STRETCH'; });
    var notfit = rows.filter(function(r){ return r[4] === 'NOT_FIT'; });

    function card(r) {
      var verdict = r[4];
      var bullets = (r[10] || '').split(' | ').map(function(b){ return '<li>'+b+'</li>'; }).join('');
      return '<div class="card ' + verdict + '">' +
        '<span class="score">' + r[3] + '%</span>' +
        '<h3>' + r[1] + ' @ ' + r[2] + '</h3>' +
        '<p><strong>Verdict:</strong> ' + verdict + ' &nbsp; <a href="' + r[12] + '" target="_blank">View Job</a></p>' +
        (r[9] ? '<p><strong>Tailored Summary:</strong> ' + r[9] + '</p>' : '') +
        (bullets ? '<p><strong>Key Bullets:</strong></p><ul>' + bullets + '</ul>' : '') +
        (r[6] ? '<p><strong>Gaps:</strong> ' + r[6].split(', ').map(function(g){ return '<span class="tag">'+g+'</span>'; }).join('') + '</p>' : '') +
        (r[8] ? '<p><strong>Why Not:</strong> ' + r[8] + '</p>' : '') +
        '<p><strong>Recommendation:</strong> ' + r[11] + '</p>' +
        '</div>';
    }

    html += '<h2>STRONG FIT (' + strong.length + ')</h2>' + (strong.map(card).join('') || '<p>None today</p>');
    html += '<h2>STRETCH (' + stretch.length + ')</h2>' + (stretch.map(card).join('') || '<p>None today</p>');
    html += '<h2>NOT FIT (' + notfit.length + ')</h2>' + (notfit.map(card).join('') || '<p>None today</p>');
  }

  html += '</body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('Jobright Scanner');
}
