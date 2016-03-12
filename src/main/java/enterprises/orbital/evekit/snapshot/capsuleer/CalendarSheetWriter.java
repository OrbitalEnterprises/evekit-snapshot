package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CalendarEventAttendee;
import enterprises.orbital.evekit.model.character.UpcomingCalendarEvent;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class CalendarSheetWriter {

  // Singleton
  private CalendarSheetWriter() {}

  public static List<Long> dumpCalendarEventAttendees(
                                                      SynchronizedEveAccount acct,
                                                      ZipOutputStream stream,
                                                      List<UpcomingCalendarEvent> events,
                                                      long at) throws IOException {
    List<Long> itemIDs = new ArrayList<Long>();
    stream.putNextEntry(new ZipEntry("CalendarEventAttendees.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Event ID", "Character ID", "Character Name", "Response");
    for (UpcomingCalendarEvent event : events) {
      int eventID = event.getEventID();
      List<CalendarEventAttendee> allAttendees = CalendarEventAttendee.getByEventID(acct, at, eventID);
      if (allAttendees.size() > 0) {
        for (CalendarEventAttendee next : allAttendees) {
          // @formatter:off
          SheetUtils.populateNextRow(output,
                                     new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getEventID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getCharacterName(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getResponse(), SheetUtils.CellFormat.NO_STYLE));
          // @formatter:on
          itemIDs.add(next.getCid());
        }
        output.println();
      }
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // UpcomingCalendarEvents.csv
    // UpcomingCalendarEventsMeta.csv
    // CalendarEventAttendees.csv
    // CalendarEventAttendeesMeta.csv
    stream.putNextEntry(new ZipEntry("UpcomingCalendarEvents.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Event ID", "Event Title", "Event Text", "Event Date (Raw)", "Event Date", "Duration", "Owner ID", "Owner Name", "Response",
                       "Important");
    List<UpcomingCalendarEvent> events = UpcomingCalendarEvent.getAllUpcomingCalendarEvents(acct, at);
    List<Long> metaIDs = new ArrayList<Long>();

    // Write out event data first
    for (UpcomingCalendarEvent next : events) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getEventID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getEventTitle(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getEventText(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getEventDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(new Date(next.getEventDate()), SheetUtils.CellFormat.DATE_STYLE), 
                                 new DumpCell(next.getDuration(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getOwnerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getOwnerName(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getResponse(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.isImportant(), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    if (events.size() > 0) {
      // Wrote at least event so proceed

      // Write out meta data for events
      output = SheetUtils.prepForMetaData("UpcomingCalendarEventsMeta.csv", stream, false, null);
      for (Long next : metaIDs) {
        int count = SheetUtils.dumpNextMetaData(acct, output, next, "UpcomingCalendarEvent");
        if (count > 0) output.println();
      }
      output.flush();
      stream.closeEntry();

      // Write out CalendarEventAttendees in the same style as meta data.
      metaIDs = dumpCalendarEventAttendees(acct, stream, events, at);
      if (metaIDs.size() > 0) {
        // Only write out meta-data if an event attendee was actually written.
        output = SheetUtils.prepForMetaData("CalendarEventAttendeesMeta.csv", stream, false, null);
        for (Long next : metaIDs) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "CalendarEventAttendee");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }
    }
  }

}
