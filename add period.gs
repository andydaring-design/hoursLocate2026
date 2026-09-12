/**
 * @OnlyCurrentDoc
 */

// ------------------------------------------------------------
// Adds another working-time period to the day selected.
//
// can select ANY CELL belonging to the day.
// ------------------------------------------------------------

function addTimePeriod() {

  // ------------------------------------------------------------
  // 1. get worksheet
  // ------------------------------------------------------------

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // selected cell
  const selectedCell = sheet.getActiveCell();

  // counts from 1
  const selectedRow = selectedCell.getRow();


  // ------------------------------------------------------------
  // 2. find start of the day
  // ------------------------------------------------------------

  // get cell's "header" row
  // walk up until we find a row with a number in column B (the day number).
  // example:
  //
  // Monday 7   08:00 12:00
  //             13:00 17:00   <- if clickl here
  //
  // it walks up and finds "7" 
  // this code might be shit - feels like there is a better way
  let dayRow = selectedRow;
  while (dayRow >= 1) {
    const dayNumber = sheet.getRange(dayRow, 2).getValue();
    if (dayNumber !== "") {
      break;
    }
    dayRow--;
  }


  // ------------------------------------------------------------
  // 3. did we find a day?
  // -----------------------------------------------------------

  // error instead of blundering on, to stop fuck ups
  if (dayRow < 1) {
    SpreadsheetApp.getUi().alert(
      "I couldn't find a day above the selected cell."
    );
    return;
  }

  // ------------------------------------------------------------
  // 4. find last row of day
  // ------------------------------------------------------------

  // start at the day row and move (robert) downy (jnr)
  // carry on till we hit the next day number
  let lastDayRow = dayRow;
  const lastSheetRow = sheet.getLastRow();
  while (lastDayRow < lastSheetRow) {
    const nextDayNumber =
      sheet.getRange(lastDayRow + 1, 2).getValue();
    if (nextDayNumber !== "") {
      break;
    }
    lastDayRow++;
  }


  // ------------------------------------------------------------
  // 5. slip in a new row 
  // ------------------------------------------------------------

  // new row after last row of the day (already addedm rows)
  sheet.insertRowAfter(lastDayRow);
  const newRow = lastDayRow + 1;


  // ------------------------------------------------------------
  // 6. copy formatting
  // ------------------------------------------------------------

  // copy the last row belonging to this day
  //
  // (this should make sure all new rows inherit the formatting, color.. blah)

  const oldRowRange =
    sheet.getRange(lastDayRow, 1, 1, 6);

  const newRowRange =
    sheet.getRange(newRow, 1, 1, 6);

  oldRowRange.copyTo(
    newRowRange,
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );


  // ------------------------------------------------------------
  // 7. copy row height (i shouldn't need to do this, but... just do it)
  // ------------------------------------------------------------

  const oldHeight =
    sheet.getRowHeight(lastDayRow);

  sheet.setRowHeight(newRow, oldHeight);


  // ------------------------------------------------------------
  // 8. clear anything from A, B, D & E
  // ------------------------------------------------------------

  // clear only the contents.
  //
  // i feel like this could be done better, I should learn more
  sheet.getRange(newRow, 1, 1, 2).clearContent();
  sheet.getRange(newRow, 4, 1, 2).clearContent();


  // ------------------------------------------------------------
  // 10. make the hours formula thingy
  // ------------------------------------------------------------

  // wouldn;t it be better to just copy the formula? huh?
  sheet
    .getRange(newRow, 6)
    .setFormula(`=E${newRow}-D${newRow}`);

  // ------------------------------------------------------------
  // 11. move cursor to D cell - for making it easy.... why do i do this?
  // ------------------------------------------------------------

  sheet.setActiveRange(
    sheet.getRange(newRow, 4)
  );


  // ------------------------------------------------------------
  // we are done! 
  // ------------------------------------------------------------
  // Copyright 1984, Winston Smith
  // just joking, but seriously if you're nicking my code you are 
  // scraping the bottom of the barrel
  // ------------------------------------------------------------

}
