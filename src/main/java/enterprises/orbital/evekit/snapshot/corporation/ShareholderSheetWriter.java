package enterprises.orbital.evekit.snapshot.corporation;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.corporation.Shareholder;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class ShareholderSheetWriter {

  // Singleton
  private ShareholderSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Shareholders.csv
    // ShareholdersMeta.csv
    stream.putNextEntry(new ZipEntry("Shareholders.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Shareholder ID", "Shareholder Type", "Shares");
    List<Long> metaIDs = new ArrayList<>();
    List<Shareholder> batch = CachedData.retrieveAll(at,
                                                     (contid, at1) -> Shareholder.accessQuery(acct, contid, 1000, false,
                                                                                              at1,
                                                                                              AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any()));

    for (Shareholder next : batch) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getShareholderID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getShareholderType(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getShares(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("ShareholdersMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "Shareholder");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
