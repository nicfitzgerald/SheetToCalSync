/**
 * Built with inspiration from https://www.youtube.com/watch?v=kYmdsOvUOng
 * Follow instructions to obtain calender ID
 * Script assumes a 3 column spreadsheet with a single row taken by headers
 * BILL NAME | DUE DATE | AMOUNT
 */


/** ID for Calendar to update */
const CALID = "CALENDER_ID_FROM_VIDEO"

/** Main function to create calendar event */
function createCalendarEvent() {
  /** 
   * Stores calendar 
   * */
  let billCalender = CalendarApp.getCalendarById(CALID);
  /** 
   * The sheet we're going to be pulling information from 
   * */
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("[Name of Sheet in Spreadsheet]");
  /** 
   * This gets all of the data, so you don't have to worry about updating with more rows 
   * */
  let bill = sheet.getDataRange().getValues();
  /** 
   * Cuts the first row off, which is designed to be headers
   * If you have more than 1 row for headers, adjust the
   * 1 accordingly (2 header rows would be 2, etc)
   */
  bill.splice(0, 1);

  /** 
   * Deletes an existing events on the calendar upon update 
   * 
   * @param evArr = Array of events
  */
  function deleteEvents(evArr) {
    let delArr = []
    let eventId
    evArr.forEach(event => {
      eventId = event.getId()
      if (delArr.includes(eventId)) {
        return
      } else {
        Logger.log(`Event ${event.getTitle()} with ID ${eventId} has been deleted.`)
        delArr.push(eventId)
        billCalender.getEventSeriesById(event.getId()).deleteEventSeries()
      }
    })
  }

  /** 
   * This gets a clean start date 
   * */
  function startDateGen(dueDate) {
    let month = new Date()
    month = month.getMonth()
    month += 1
    let day = dueDate
    let year = new Date()
    year = year.getFullYear()
    return new Date(`${month} ${day}, ${year}`)
  }

  /**
   * This is what creates the event
   * 
   * @param entry
   *        the spreadsheet data
   * 
   * @param endDate
   *        a specific date (in case of a lease, car payment, etc) or arbitrary
   *        DEFAULT is END_DATE
   * @param color
   *        what color you want the event to be
   *        https://developers.google.com/apps-script/reference/calendar/event-color
   *        (defaults to 5, which is yellow)
   * @returns the created event
   */
  function createEv(entry, endDate = END_DATE, color = 5) {
    let expense = billCalender.createAllDayEventSeries(
      entry[0],
      startDateGen(entry[1]),
      CalendarApp
        .newRecurrence()
        .addMonthlyRule()
        .until(new Date(endDate)),
      {
        description: `Total Due: $${entry[2].toString().replace(/-/gm, '')}`
      }
    )
    expense.setColor(color)
    Logger.log(`Bill ${expense.getTitle()} has been created.`)
    return expense
  }

  /**
   * Arbritary start date
   * 1/1/2022
   * Change it to whatever you see fit
   */
  const START_DATE = new Date(2022, 1, 1, 0, 0, 0)
  /**
   * Arbitrary end date
   * 12/31/2027
   * Change it to whatever you see fit
   */
  const END_DATE = new Date(2027, 12, 31, 0, 0, 0)

  /** 
   * Gets the array of events passed into the event deletion function
   */
  getEv = billCalender.getEvents(START_DATE, END_DATE)
  if (getEv.length > 0) {
    deleteEvents(getEv)
  }

  /**
   * Loops loops over the spreadsheet data
   * 
   * `entry` is an array, ['Bill Name', 'Due Date', 'Amount']
   */
  bill.forEach(function(entry) {
    /** 
     * Switch case to single out individual bills for different processing
     * 
     * Use cases:
     * switch(entry[0]) {
     *  case 'Rent':
     *    createEv(entry, 'December 28, 2022', 1)
     *    break
     *  default:
     *    createEv(entry)
     *    break
     * }
     */
    switch (entry[0]) {
      case '[CASE NAME]':
        createEv(entry, '[END DATE]', ['COLOR CODE FROM LINK'])
        break
      default:
        createEv(entry)
        break
    }
  });
}