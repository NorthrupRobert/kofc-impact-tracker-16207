function buildMemberLevelTable() {
  const ss = SpreadsheetApp.getActive();
  const eventsSheet = ss.getSheetByName("Events");
  const outputSheet = ss.getSheetByName("Member Analysis (Auto)") || ss.insertSheet("Member Analysis (Auto)");

  // Clear old output
  outputSheet.clearContents();

  // Get all event data
  const data = eventsSheet.getRange(
    2,
    1,
    eventsSheet.getLastRow() - 1,
    eventsSheet.getLastColumn()
  ).getValues();

  // Identify column indexes
  const header = eventsSheet.getRange(1, 1, 1, eventsSheet.getLastColumn()).getValues()[0];

  const eventIdIndex = header.indexOf("Event_ID");
  const membersIndex = header.indexOf("Names_of_Participating_Members");
  const eventDateIndex = header.indexOf("End_Date");
  const eventHoursIndex = header.indexOf("Total_Event_Hours");
  const eventTypeIndex = header.indexOf("Volunteer_(V)_Faternal_(Ft)_Faith_(Fi)_or_Meeting_(M)_Event");
  const eventProgramIndex = header.indexOf("Program");

  if (eventIdIndex === -1 || membersIndex === -1) {
    throw new Error("Could not find Event_ID or Names_of_Participating_Members columns.");
  }

  // Build the member-level table header
  const output = [[
    "Event_ID",
    "Member",
    "Date",
    "Total_Event_Hours",
    "Event_Type",
    "Program",
    "First_Event_Date",
    "First_Event_Month",
    "Activated"
  ]];

  // Temporary storage for raw rows (before activation logic)
  const rawRows = [];

  // Pass 1: Build raw rows
  data.forEach(row => {
    const eventId = row[eventIdIndex];
    const membersRaw = row[membersIndex];

    if (!eventId || !membersRaw) return;

    const eventDate = row[eventDateIndex];
    const eventHours = row[eventHoursIndex];
    const eventType = row[eventTypeIndex];
    const eventProgram = row[eventProgramIndex];

    const members = membersRaw
      .split(",")
      .map(m => m.trim())
      .filter(m => m.length > 0);

    members.forEach(member => {
      rawRows.push({
        eventId,
        member,
        eventDate,
        eventHours,
        eventType,
        eventProgram
      });
    });
  });

  // Pass 2: Compute earliest event per member
  const earliest = {};

  rawRows.forEach(r => {
    if (!earliest[r.member] || r.eventDate < earliest[r.member]) {
      earliest[r.member] = r.eventDate;
    }
  });

  // Pass 3: Add activation metadata and push to output
  rawRows.forEach(r => {
    const firstDate = earliest[r.member];
    const firstMonth = Utilities.formatDate(firstDate, Session.getScriptTimeZone(), "yyyy-MM");

    const activated = (r.eventDate.getTime() === firstDate.getTime()) ? "TRUE" : "";

    output.push([
      r.eventId,
      r.member,
      r.eventDate,
      r.eventHours,
      r.eventType,
      r.eventProgram,
      firstDate,
      firstMonth,
      activated
    ]);
  });

  // Write output
  outputSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Analytics")
    .addItem("Refresh Member Table", "buildMemberLevelTable")
    .addToUi();
}
