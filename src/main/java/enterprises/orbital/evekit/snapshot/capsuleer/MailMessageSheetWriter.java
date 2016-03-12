package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterMailMessage;
import enterprises.orbital.evekit.model.character.CharacterMailMessageBody;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class MailMessageSheetWriter {

  // Singleton
  private MailMessageSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // MailMessages.csv
    // MailMessagesMeta.csv
    stream.putNextEntry(new ZipEntry("MailMessages.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Message ID", "Sender ID", "Sender Name", "ToCharacter IDs", "Sent Date (Raw)", "Sent Date", "Title", "CorpOrAlliance ID",
                       "ToList IDs", "Msg Read", "SenderTypeID", "Body Retrieved", "Body");
    List<Long> metaIDs = new ArrayList<Long>();
    long contid = -1;
    List<Long> batch = CharacterMailMessage.getMessageIDs(acct, at, false, 1000, contid);
    while (batch.size() > 0) {
      for (Long next : batch) {
        CharacterMailMessage nextMsg = CharacterMailMessage.get(acct, at, next);
        CharacterMailMessageBody body = CharacterMailMessageBody.get(acct, at, next);
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(nextMsg.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(nextMsg.getMessageID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(nextMsg.getSenderID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(nextMsg.getSenderName(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(Arrays.toString(nextMsg.getToCharacterID().toArray(new Long[nextMsg.getToCharacterID().size()])), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(nextMsg.getSentDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(nextMsg.getSentDate()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(nextMsg.getTitle(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(nextMsg.getCorpOrAllianceID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(Arrays.toString(nextMsg.getToListID().toArray(new Long[nextMsg.getToListID().size()])), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(nextMsg.isMsgRead(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(nextMsg.getSenderTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(body.isRetrieved(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(body.getBody(), SheetUtils.CellFormat.NO_STYLE)); 
        // @formatter:on
        metaIDs.add(nextMsg.getCid());
        contid = nextMsg.getSentDate();
      }
      batch = CharacterMailMessage.getMessageIDs(acct, at, false, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MailMessagesMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CharacterMailMessage");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
