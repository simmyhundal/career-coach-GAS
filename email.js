function escapeHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildTaskPlanHtml(aiContent) {
  const TIMED_CARD = [
    'border-left: 3px solid #2D7DD2;',
    'padding: 12px 16px;',
    'margin-bottom: 10px;',
    'background-color: #f5f9ff;',
    'border-radius: 0 6px 6px 0;'
  ].join(' ');

  const DAILY_CARD = [
    'border-left: 3px solid #1E8A5E;',
    'padding: 12px 16px;',
    'margin-bottom: 10px;',
    'background-color: #f3fbf7;',
    'border-radius: 0 6px 6px 0;'
  ].join(' ');

  const DAILY_BADGE = [
    'display: inline-block;',
    'font-size: 10px;',
    'font-weight: 700;',
    'text-transform: uppercase;',
    'letter-spacing: 0.5px;',
    'background-color: #1E8A5E;',
    'color: #ffffff;',
    'padding: 2px 7px;',
    'border-radius: 10px;',
    'margin-left: 8px;',
    'vertical-align: middle;'
  ].join(' ');

  const lines = (aiContent || "").split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const html = [];
  let inList = false;

  lines.forEach(line => {
    const headingMatch = line.match(/^###\s+(.*)$/);
    const bulletMatch = line.match(/^[-*]\s+(.*)$/);

    if (headingMatch) {
      if (inList) { html.push("</ul>"); inList = false; }
      const content = headingMatch[1].trim();
      const timedMatch = content.match(/^(.+?):\s*(\d+)\s+mins?$/i);

      if (timedMatch) {
        const name = escapeHtml(timedMatch[1].trim());
        const mins = timedMatch[2];
        html.push(
          `<div style="${TIMED_CARD}">` +
          `<div style="font-size: 14px; font-weight: 600; color: #1A2744;">${name}</div>` +
          `<div style="font-size: 12px; color: #8695A3; margin-top: 4px;">${mins} minutes</div>` +
          `</div>`
        );
      } else {
        const name = escapeHtml(content);
        html.push(
          `<div style="${DAILY_CARD}">` +
          `<div style="font-size: 14px; font-weight: 600; color: #1A2744;">${name}` +
          `<span style="${DAILY_BADGE}">Daily</span></div>` +
          `</div>`
        );
      }
    } else if (bulletMatch) {
      if (!inList) {
        html.push('<ul style="margin: 8px 0 8px 4px; padding-left: 18px; color: #5D6D7E; font-size: 13px;">');
        inList = true;
      }
      html.push(`<li style="margin-bottom: 4px;">${escapeHtml(bulletMatch[1])}</li>`);
    } else {
      if (inList) { html.push("</ul>"); inList = false; }
      html.push(`<p style="font-size: 13px; color: #5D6D7E; margin: 6px 0;">${escapeHtml(line)}</p>`);
    }
  });

  if (inList) html.push("</ul>");

  return html.join("\n") ||
    '<p style="font-size: 13px; color: #8695A3; font-style: italic;">No tasks available today.</p>';
}

function markdownToPlainText(markdown) {
  return (markdown || "")
    .replace(/^###\s+/gm, "")
    .replace(/\*\*(.+?)\*\*/g, "$1")
    .replace(/^[-*]\s+/gm, "- ");
}

function buildUpcomingEventsPlainText(events) {
  if (!events || !events.length) {
    return "No calendar events scheduled during your workday.";
  }

  return events.map(event => `- ${event.timeLabel}: ${event.title}`).join("\n");
}

function buildUpcomingEventsHtml(events) {
  if (!events || !events.length) {
    return '<p style="font-size: 13px; color: #8695A3; font-style: italic;">No calendar events during your workday.</p>';
  }

  return events.map((event, i) => {
    const border = i < events.length - 1 ? 'border-bottom: 1px solid #f0f3f6;' : '';
    return (
      `<div style="padding: 9px 0; ${border}">` +
      `<table style="border-collapse: collapse; width: 100%;"><tr>` +
      `<td style="padding: 0; width: 160px; vertical-align: top; white-space: nowrap;">` +
      `<span style="font-size: 13px; color: #2D7DD2; font-weight: 500;">${escapeHtml(event.timeLabel)}</span>` +
      `</td>` +
      `<td style="padding: 0; vertical-align: top;">` +
      `<span style="font-size: 13px; color: #1A2744;">${escapeHtml(event.title)}</span>` +
      `</td>` +
      `</tr></table>` +
      `</div>`
    );
  }).join("\n");
}

function buildFeaturedJobsPlainText(jobs) {
  if (!jobs || !jobs.length) {
    return "No recent Airtable jobs were found today.";
  }

  return jobs.map((job, index) => {
    const parts = [
      `${index + 1}. ${job.title || "Untitled role"}`,
      job.targetDate ? `Apply by: ${formatJobDate(job.targetDate)}` : "",
      job.url ? `Link: ${job.url}` : ""
    ].filter(Boolean);

    return parts.join("\n");
  }).join("\n\n");
}

function formatJobDate(isoStr) {
  if (!isoStr) return "";
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const parts = isoStr.split("-");
  if (parts.length !== 3) return isoStr;
  return `${months[parseInt(parts[1], 10) - 1]} ${parseInt(parts[2], 10)}, ${parts[0]}`;
}

function buildFeaturedJobsHtml(jobs) {
  if (!jobs || !jobs.length) {
    return '<p style="font-size: 13px; color: #8695A3; font-style: italic;">No recent job opportunities found.</p>';
  }

  return jobs.map((job, i) => {
    const border = i < jobs.length - 1 ? 'border-bottom: 1px solid #f0f3f6;' : '';
    const title = escapeHtml(job.title || "Untitled role");
    const targetDatePart = job.targetDate
      ? `<span style="color: #8695A3; font-size: 12px; margin-right: 10px;">Apply by: ${escapeHtml(formatJobDate(job.targetDate))}</span>`
      : "";
    const linkPart = job.url
      ? `<a href="${escapeHtml(job.url)}" style="color: #2D7DD2; font-size: 12px; text-decoration: none;">View posting &rarr;</a>`
      : "";
    const meta = targetDatePart || linkPart
      ? `<div style="margin-top: 5px;">${targetDatePart}${linkPart}</div>`
      : "";

    return (
      `<div style="padding: 12px 0; ${border}">` +
      `<div style="font-size: 14px; font-weight: 600; color: #1A2744; margin-bottom: 3px;">${title}</div>` +
      meta +
      `</div>`
    );
  }).join("\n");
}

function buildEmailHtml(dateLabel, availability, aiContent, featuredJobs, okrUrl) {
  const S = {
    wrap: 'font-family: "Helvetica Neue", Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f5f7fa;',
    header: 'background-color: #1E3A5F; padding: 28px 32px 20px;',
    eyebrow: 'color: rgba(255,255,255,0.55); font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 8px;',
    title: 'color: #ffffff; font-size: 26px; font-weight: 700; line-height: 1.2; margin-bottom: 4px;',
    date: 'color: rgba(255,255,255,0.65); font-size: 14px;',
    statsBar: 'background-color: #16304F; padding: 14px 32px;',
    statsTable: 'border-collapse: collapse; width: 100%;',
    statsCell: 'color: #ffffff; padding: 0; width: 50%;',
    statsNum: 'font-size: 18px; font-weight: 700;',
    statsLabel: 'font-size: 12px; color: rgba(255,255,255,0.6); margin-left: 5px;',
    body: 'padding: 20px 24px; background-color: #f5f7fa;',
    card: 'background-color: #ffffff; border-radius: 8px; padding: 20px 24px; margin-bottom: 16px; border: 1px solid #e2e8ef;',
    sectionLabel: 'font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; color: #8695A3; margin-bottom: 14px;',
    tagline: 'text-align: center; padding: 6px 0 4px; color: #8695A3; font-size: 13px; font-style: italic;',
    footer: 'background-color: #1E3A5F; padding: 14px 32px; text-align: center;',
    footerText: 'color: rgba(255,255,255,0.4); font-size: 11px;'
  };

  return [
    `<div style="${S.wrap}">`,

    `<div style="${S.header}">`,
    `<div style="${S.eyebrow}">AI Career Coach</div>`,
    `<div style="${S.date}">${escapeHtml(dateLabel)}</div>`,
    `</div>`,

    `<div style="${S.statsBar}">`,
    `<table style="${S.statsTable}"><tr>`,
    `<td style="${S.statsCell}">`,
    `<span style="${S.statsNum}">${availability.totalMinutes}</span>`,
    `<span style="${S.statsLabel}">min free today</span>`,
    `</td>`,
    `<td style="${S.statsCell}">`,
    `<span style="${S.statsNum}">${availability.largestBlock}</span>`,
    `<span style="${S.statsLabel}">min focus block</span>`,
    `</td>`,
    `</tr></table>`,
    `</div>`,

    `<div style="${S.body}">`,

    `<div style="${S.card}">`,
    `<div style="${S.sectionLabel}">Today&#39;s Schedule</div>`,
    buildUpcomingEventsHtml(availability.upcomingEvents),
    `</div>`,

    `<div style="${S.card}">`,
    `<div style="${S.sectionLabel}">Priority Tasks</div>`,
    buildTaskPlanHtml(aiContent),
    `</div>`,

    `<div style="${S.card}">`,
    `<div style="${S.sectionLabel}">Featured PM Roles</div>`,
    buildFeaturedJobsHtml(featuredJobs),
    `</div>`,

    `<div style="${S.tagline}">Go get &#39;em!</div>`,

    `</div>`,

    `<div style="${S.footer}">`,
    `<div style="${S.footerText}">AI Career Coach &middot; Google Apps Script</div>`,
    okrUrl
      ? `<div style="margin-top: 6px;"><a href="${escapeHtml(okrUrl)}" style="color: rgba(255,255,255,0.55); font-size: 11px; text-decoration: none;">📊 View OKR Sheet</a></div>`
      : "",
    `</div>`,

    `</div>`
  ].join("\n");
}

function sendCoachEmail(recipient, aiContent, availability, featuredJobs, okrUrl) {
  const dateLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");
  const subject = `Daily Career Coach: ${dateLabel}`;
  const aiPlain = markdownToPlainText(aiContent).trim();
  const upcomingEventsPlain = buildUpcomingEventsPlainText(availability.upcomingEvents);
  const featuredJobsPlain = buildFeaturedJobsPlainText(featuredJobs);

  const plainBody = [
    "Good morning!",
    "",
    `Today you have ${availability.totalMinutes} minutes of White Space across your calendars.`,
    "",
    "Coming up on your calendar:",
    upcomingEventsPlain,
    "",
    "The following are your proposed tasks:",
    "",
    aiPlain,
    "",
    "Most recently created jobs from Airtable:",
    "",
    featuredJobsPlain,
    "",
    okrUrl ? `OKR Sheet: ${okrUrl}` : "",
    "",
    "Go get 'em!"
  ].filter(line => line !== undefined).join("\n");

  GmailApp.sendEmail(recipient, subject, plainBody, {
    htmlBody: buildEmailHtml(dateLabel, availability, aiContent, featuredJobs, okrUrl)
  });
}
