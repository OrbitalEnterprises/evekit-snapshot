package enterprises.orbital.evekit.snapshot.corporation;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.corporation.CorporationMedal;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class CorporationMedalSheetWriter {

  // Singleton
  private CorporationMedalSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // CorporationMedals.csv
    // CorporationMedalsMeta.csv
    stream.putNextEntry(new ZipEntry("CorporationMedals.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Medal ID", "Title", "Description", "Created (Raw)", "Created", "Creator ID");
    List<Long> metaIDs = new ArrayList<Long>();
    List<CorporationMedal> medals = CorporationMedal.getAll(acct, at);
    for (CorporationMedal next : medals) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getMedalID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTitle(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getDescription(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getCreated(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(new Date(next.getCreated()), SheetUtils.CellFormat.DATE_STYLE), 
                                 new DumpCell(next.getCreatorID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("CorporationMedalsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CorporationMedal");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
