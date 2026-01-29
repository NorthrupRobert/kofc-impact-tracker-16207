function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActive();
  const eventsSheet = ss.getSheetByName("Events");

  // Raw form submission values in order
  const row = e.values;

  // FORM COLUMN ORDER:
  //
  // 0: Timestamp
  // 1: Event Title
  // 2: Event Type (V/Ft/Fi/M)
  // 3: Event Program
  // 4: Start Date
  // 5: End Date
  // 6: Organizer
  // 7: Names of Participating Members (comma-separated EXACTLY)
  // 8: Number of Participating Non-Members
  // 9: Event Hours (per-member average)
  // 10: Cost
  // 11: Earnings
  // 12: Units Sold or Delivered
  // 13: Individuals Serviced
  // 14: Impact
  // 15: Notes

  // EVENTS SHEET COLUMN ORDER (your structure):
  //
  // 1.  Start_Date
  // 2.  End_Date
  // 3.  Event_Name
  // 4.  Event_ID                      (leave blank → formula fills)
  // 5.  Program
  // 6.  Organizer
  // 7.  Names_of_Participating_Members
  // 8.  Unknown_Members               (leave blank)
  // 9.  Number_of_Participating_Members (leave blank → formula)
  // 10. Number_of_Participating_Non-Members
  // 11. Total_Participants            (leave blank → formula)
  // 12. Volunteer_(V)_Fraternal_(Ft)_Faith_(Fi)_or_Meeting_(M)_Event
  // 13. Total_Event_Hours
  // 14. Volunteer_Hours_Acquired      (leave blank → formula)
  // 15. Fraternal_Hours_Acquired      (leave blank → formula)
  // 16. Faith_Hours_Acquired          (leave blank → formula)
  // 17. Meeting_Hours_Acquired        (leave blank → formula)
  // 18. Cost
  // 19. Earnings
  // 20. Budget_Affect                 (leave blank → formula)
  // 21. Units_Sold_or_Delivered
  // 22. Cost_Per_Unit                 (leave blank → formula)
  // 23. Individuals_Serviced
  // 24. Impact
  // 25. Notes

  const formattedRow = [
    row[4],   // Start_Date
    row[5],   // End_Date
    row[1],   // Event_Name
    "",       // Event_ID (formula-generated)
    row[3],   // Program
    row[6],   // Organizer
    row[7],   // Names_of_Participating_Members (DO NOT ALTER FORMAT)
    "",       // Unknown_Members
    "",       // Number_of_Participating_Members (formula)
    row[8],   // Number_of_Participating_Non-Members
    "",       // Total_Participants (formula)
    row[2],   // Event Type (V/Ft/Fi/M)
    row[9],   // Total_Event_Hours
    "",       // Volunteer_Hours_Acquired (formula)
    "",       // Fraternal_Hours_Acquired (formula)
    "",       // Faith_Hours_Acquired (formula)
    "",       // Meeting_Hours_Acquired (formula)
    row[10],  // Cost
    row[11],  // Earnings
    "",       // Budget_Affect (formula)
    row[12],  // Units_Sold_or_Delivered
    "",       // Cost_Per_Unit (formula)
    row[13],  // Individuals_Serviced
    row[14],  // Impact
    row[15]   // Notes
  ];

  // Append the row to the Events sheet
  eventsSheet.appendRow(formattedRow);
}