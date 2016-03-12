package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterContactNotification;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class ContactNotificationSheetWriter {

  // Singleton
  private ContactNotificationSheetWriter() {}

  public static final Comparator<CharacterContactNotification> ascendingNotificationComparator = new Comparator<CharacterContactNotification>() {

    @Override
    public int compare(
                       CharacterContactNotification o1,
                       CharacterContactNotification o2) {
      if (o1.getSentDate() < o2.getSentDate()) return -1;
      if (o1.getSentDate() == o2.getSentDate()) return 0;
      return 1;
    }

  };

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // ContactNotifications.csv
    // ContactNotificationsMeta.csv
    stream.putNextEntry(new ZipEntry("ContactNotifications.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Notification ID", "Sender ID", "SenderName", "Sent Date (Raw)", "Sent Date", "Message Data");
    List<Long> metaIDs = new ArrayList<Long>();
    long contid = -1;
    List<CharacterContactNotification> batch = CharacterContactNotification.getAllNotifications(acct, at, 1000, contid);
    while (batch.size() > 0) {
      for (CharacterContactNotification next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getNotificationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getSenderID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getSenderName(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getSentDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(next.getSentDate()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getMessageData(), SheetUtils.CellFormat.NO_STYLE)); 
        // @formatter:on
        metaIDs.add(next.getCid());
      }
      contid = batch.get(batch.size() - 1).getSentDate();
      batch = CharacterContactNotification.getAllNotifications(acct, at, 1000, contid);
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
