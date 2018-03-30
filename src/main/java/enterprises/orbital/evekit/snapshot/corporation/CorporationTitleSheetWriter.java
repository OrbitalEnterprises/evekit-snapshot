package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.CorporationTitle;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class CorporationTitleSheetWriter {

  // Singleton
  private CorporationTitleSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // CorporationTitles.csv
    // CorporationTitlesMeta.csv
    stream.putNextEntry(new ZipEntry("CorporationTitles.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Title ID", "Title Name");
    List<Long> metaIDs = new ArrayList<>();
    List<CorporationTitle> batch = CachedData.retrieveAll(at,
                                                          (contid, at1) -> CorporationTitle.accessQuery(acct, contid,
                                                                                                        1000, false,
                                                                                                        at1,
                                                                                                        AttributeSelector.any(),
                                                                                                        AttributeSelector.any()));

    for (CorporationTitle next : batch) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getTitleID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTitleName(), SheetUtils.CellFormat.NO_STYLE));
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("CorporationTitlesMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CorporationTitle");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
