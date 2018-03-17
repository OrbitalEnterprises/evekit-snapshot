package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.CharacterContactNotification;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ContactNotificationSheetWriter {

  // Singleton
  private ContactNotificationSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // ContactNotifications.csv
    // ContactNotificationsMeta.csv
    stream.putNextEntry(new ZipEntry("ContactNotifications.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Notification ID", "Sender ID", "Sent Date (Raw)", "Sent Date", "Standing Level",
                       "Message Data");
    List<Long> metaIDs = new ArrayList<>();
    List<CharacterContactNotification> batch = CachedData.retrieveAll(at,
                                                                      (contid, at1) -> CharacterContactNotification.accessQuery(
                                                                          acct, contid, 1000, false,
                                                                          at1,
                                                                          AttributeSelector.any(),
                                                                          AttributeSelector.any(),
                                                                          AttributeSelector.any(),
                                                                          AttributeSelector.any(),
                                                                          AttributeSelector.any()));
    for (CharacterContactNotification next : batch) {
      // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getNotificationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getSenderID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getSentDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getSentDate()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getStandingLevel(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                   new DumpCell(next.getMessageData(), SheetUtils.CellFormat.NO_STYLE));
        // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("ContactNotificationsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CharacterContactNotification");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
