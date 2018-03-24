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
import enterprises.orbital.evekit.model.corporation.Division;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class CorporationDivisionSheetWriter {

  // Singleton
  private CorporationDivisionSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Divisions.csv
    // DivisionsMeta.csv
    stream.putNextEntry(new ZipEntry("Divisions.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Division", "Name", "Wallet");
    List<Long> metaIDs = new ArrayList<>();
    List<Division> batch = CachedData.retrieveAll(at,
                                                  (contid, at1) -> Division.accessQuery(acct, contid, 1000, false, at1,
                                                                                        AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any()));

    for (Division next : batch) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getDivision(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getName(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.isWallet(), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("DivisionsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "Division");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
