package enterprises.orbital.evekit.snapshot.common;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.common.ContactLabel;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ContactLabelSheetWriter {

  // Singleton
  private ContactLabelSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // ContactLabels.csv
    // ContactLabelsMeta.csv
    stream.putNextEntry(new ZipEntry("ContactLabels.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "List", "Label ID", "Name");
    List<ContactLabel> labels = CachedData.retrieveAll(at, (contid, at1) -> ContactLabel.accessQuery(acct, contid, 1000,
                                                                                                     false,
                                                                                                     at1,
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any()));

    for (ContactLabel next : labels) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getList(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getLabelID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getName(), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("ContactLabelsMeta.csv", stream, false, null);
    for (ContactLabel next : labels) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "ContactLabel");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
