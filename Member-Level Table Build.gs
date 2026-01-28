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
    "End_Date",
    "Total_Event_Hours",
    "Event_Type",
    "Program"
  ]];

  // Build rows
  data.forEach(row => {
    const eventId = row[eventIdIndex];
    const membersRaw = row[membersIndex];

    if (!eventId || !membersRaw) return;

    const eventDate = row[eventDateIndex];
    const eventHours = row[eventHoursIndex];
    const eventType = row[eventTypeIndex];
    const eventProgram = row[eventProgramIndex];

    // Split members by comma, normalize spacing
    const members = membersRaw
      .split(",")
      .map(m => m.trim())
      .filter(m => m.length > 0);

    members.forEach(member => {
      output.push([
        eventId,
        member,
        eventDate,
        eventHours,
        eventType,
        eventProgram
      ]);
    });
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
