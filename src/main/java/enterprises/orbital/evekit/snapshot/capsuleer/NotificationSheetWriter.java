package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.CharacterNotification;
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
    final CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Notification ID", "Type", "Sender ID", "Sender Type", "Sent Date (Raw)", "Sent Date", "Msg Read", "Text");
    List<Long> metaIDs = new ArrayList<>();
    CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
    CachedData.stream(at, (long contid, AttributeSelector ats) ->
                          CharacterNotification.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                                           AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                           AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any()),
                      true, capture)
              .forEach(nextNote -> {
                try {
                  // @formatter:off
                  SheetUtils.populateNextRow(output,
                                   new DumpCell(nextNote.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(nextNote.getNotificationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(nextNote.getType(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(nextNote.getSenderID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(nextNote.getSenderType(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(nextNote.getSentDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(nextNote.getSentDate()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(nextNote.isMsgRead(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(nextNote.getText(), SheetUtils.CellFormat.NO_STYLE));
                  // @formatter:on
                } catch (IOException e) {
                  capture.handle(e);
                }
                metaIDs.add(nextNote.getCid());
              });

    output.flush();
    stream.closeEntry();

    // Handle MetaData
    CSVPrinter metaOutput = SheetUtils.prepForMetaData("NotificationsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, metaOutput, next, "CharacterNotification");
      if (count > 0) metaOutput.println();
    }
    metaOutput.flush();
    stream.closeEntry();
  }

}
