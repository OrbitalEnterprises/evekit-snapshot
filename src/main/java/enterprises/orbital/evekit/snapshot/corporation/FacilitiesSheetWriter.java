package enterprises.orbital.evekit.snapshot.corporation;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.corporation.Facility;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class FacilitiesSheetWriter {

  // Singleton
  private FacilitiesSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Facilities.csv
    // FacilitiesMeta.csv
    stream.putNextEntry(new ZipEntry("Facilities.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Facility ID", "Type ID", "Type Name", "Solar System ID", "Solar System Name", "Region ID", "Region Name", "Starbase Modifier",
                       "Tax");
    List<Long> metaIDs = new ArrayList<Long>();
    List<Facility> facilities = Facility.getAll(acct, at);
    for (Facility next : facilities) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getFacilityID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTypeName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getSolarSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getSolarSystemName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getRegionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getRegionName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getStarbaseModifier(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTax(), SheetUtils.CellFormat.DOUBLE_STYLE)); 
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("FacilitiesMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "Facility");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
