/**
 * 09_teamAbsences.gs
 * Team Absence Management Module
 * Displays absence records visually and allows adding/editing entries.
 */

function showTeamAbsencePlanner() {
  const html = HtmlService.createHtmlOutputFromFile('09_teamAbsencePlanner.html')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '🗓️ Team Absence Planner');
}

function getTeamAbsenceData() {
  const file = DriveApp.getFileById(CONFIG.TEAM_ABSENCE_FILE_ID);
  const json = file.getBlob().getDataAsString();
  return JSON.parse(json);
}

function saveTeamAbsenceData(data) {
  const file = DriveApp.getFileById(CONFIG.TEAM_ABSENCE_FILE_ID);
  file.setContent(JSON.stringify(data, null, 2));
}

function getCurrentTeamMembers() {
  const teamEmails = [
    'jan.fish@village-hotels.com',
    'dan.nixdorf@village-hotels.com',
    'rachel.hill@village-hotels.com'
  ];
  return getAllEmployees().filter(emp =>
    teamEmails.includes(emp.Email.toLowerCase())
  ).map(emp => ({
    name: emp.FirstName + ' ' + emp.Surname,
    email: emp.Email,
    id: emp.EmployeeNumber
  }));
}

function addNewAbsence(entry) {
  try {
    const fileId = CONFIG.TEAM_ABSENCE_FILE_ID;
    if (!fileId) throw new Error('TEAM_ABSENCE_FILE_ID is missing from CONFIG.');

    const existing = getTeamAbsenceData(); // This may also fail
    existing.push(entry);
    saveTeamAbsenceData(existing);

    Logger.log(`Added new absence for ${entry.name}: ${entry.startDate} to ${entry.endDate}`);
    return true;

  } catch (err) {
    Logger.log(`❌ Error in addNewAbsence: ${err.message}`);
    throw err;
  }
}


function detectAbsenceClashes(data) {
  const events = data.map(r => ({
    ...r,
    start: new Date(r.startDate),
    end: new Date(r.endDate)
  }));

  const clashes = [];
  for (let i = 0; i < events.length; i++) {
    for (let j = i + 1; j < events.length; j++) {
      if (events[i].type === 'WFH' || events[j].type === 'WFH') continue;
      if (
        events[i].start <= events[j].end &&
        events[i].end >= events[j].start
      ) {
        clashes.push({ a: events[i], b: events[j] });
      }
    }
  }

  return clashes;
}
