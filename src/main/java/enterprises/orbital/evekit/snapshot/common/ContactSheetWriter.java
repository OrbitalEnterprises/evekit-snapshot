package enterprises.orbital.evekit.snapshot.common;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.common.Contact;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

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
    output.printRecord("ID", "List", "Contact ID", "Contact Name", "Standing", "Contact Type ID", "In Watch List", "Label Mask");
    List<Contact> contacts = new ArrayList<Contact>();
    long contid = -1;
    List<Contact> batch = Contact.getAllContacts(acct, at, 1000, contid);
    while (batch.size() > 0) {
      contacts.addAll(batch);
      contid = contacts.get(contacts.size() - 1).getCid();
      batch = Contact.getAllContacts(acct, at, 1000, contid);
    }

    for (Contact next : contacts) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getList(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getContactID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getContactName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getStanding(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getContactTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.isInWatchlist(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getLabelMask(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
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
