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
import enterprises.orbital.evekit.model.character.CharacterNotification;
import enterprises.orbital.evekit.model.character.CharacterNotificationBody;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class NotificationSheetWriter {

  // Singleton
  private NotificationSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Notifications.csv
    // NotificationsMeta.csv
    stream.putNextEntry(new ZipEntry("Notifications.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Notification ID", "Sender ID", "Type ID", "Sent Date (Raw)", "Sent Date", "Msg Read", "Text Retrieved", "Missing", "Text");
    List<Long> metaIDs = new ArrayList<Long>();
    long contid = -1;
    List<Long> batch = CharacterNotification.getNotificationIDs(acct, at, false, 1000, contid);
    while (batch.size() > 0) {
      for (Long next : batch) {
        CharacterNotification nextNote = CharacterNotification.get(acct, at, next);
        CharacterNotificationBody body = CharacterNotificationBody.get(acct, at, next);
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(nextNote.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(nextNote.getNotificationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(nextNote.getSenderID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(nextNote.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(nextNote.getSentDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(nextNote.getSentDate()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(nextNote.isMsgRead(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(body.isRetrieved(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(body.isMissing(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(body.getText(), SheetUtils.CellFormat.NO_STYLE)); 
        // @formatter:on
        metaIDs.add(nextNote.getCid());
        contid = nextNote.getSentDate();
      }
      batch = CharacterNotification.getNotificationIDs(acct, at, false, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("NotificationsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CharacterNotification");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
