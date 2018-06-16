package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.CharacterMailMessage;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class MailMessageSheetWriter {

  // Singleton
  private MailMessageSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at)
      throws IOException {
    // Sections:
    // MailMessages.csv
    // MailMessagesMeta.csv
    stream.putNextEntry(new ZipEntry("MailMessages.csv"));
    final CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Message ID", "Sender ID", "Sent Date (Raw)", "Sent Date", "Title", "Msg Read",
                       "Labels", "Recipients", "Body");
    final List<Long> metaIDs = new ArrayList<>();
    CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
    CachedData.stream(at, (long contid, AttributeSelector ats) ->
                          CharacterMailMessage.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                                           AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                           AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                           AttributeSelector.any(), AttributeSelector.any()),
                      true, capture)
              .forEach(nextMsg -> {
                try {
                  // @formatter:off
                  SheetUtils.populateNextRow(output,
                                   new DumpCell(nextMsg.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(nextMsg.getMessageID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(nextMsg.getSenderID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(nextMsg.getSentDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(nextMsg.getSentDate()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(nextMsg.getTitle(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(nextMsg.isMsgRead(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(Arrays.toString(nextMsg.getLabels().toArray(new Integer[0])), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(Arrays.toString(nextMsg.getRecipients()
                                                                       .stream()
                                                                       .map(x -> "[" + x.getRecipientType() + ", " + x.getRecipientID() + "]")
                                                                       .toArray(String[]::new)), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(nextMsg.getBody(), SheetUtils.CellFormat.NO_STYLE));
                  // @formatter:on
                } catch (IOException e) {
                  capture.handle(e);
                }
                metaIDs.add(nextMsg.getCid());
              });

    output.flush();
    stream.closeEntry();

    // Handle MetaData
    CSVPrinter metaOutput = SheetUtils.prepForMetaData("MailMessagesMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, metaOutput, next, "CharacterMailMessage");
      if (count > 0) metaOutput.println();
    }
    metaOutput.flush();
    stream.closeEntry();
  }

}
