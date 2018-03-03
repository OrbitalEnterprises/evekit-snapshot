package enterprises.orbital.evekit.snapshot.common;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.common.Contact;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ContactSheetWriter {

  // Singleton
  private ContactSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Contacts.csv
    // ContactsMeta.csv
    stream.putNextEntry(new ZipEntry("Contacts.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "List", "Contact ID", "Standing", "Contact Type", "In Watch List", "Is Blocked", "Label ID");
    List<Contact> contacts = CachedData.retrieveAll(at,
                                                    (contid, at1) -> Contact.accessQuery(acct, contid, 1000, false, at1, AttributeSelector.any(),
                                                                                         AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                                                         AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any()));

    for (Contact next : contacts) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getList(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getContactID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getStanding(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                 new DumpCell(next.getContactType(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.isInWatchlist(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.isBlocked(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getLabelID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("ContactsMeta.csv", stream, false, null);
    for (Contact next : contacts) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "Contact");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
