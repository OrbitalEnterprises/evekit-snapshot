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
import enterprises.orbital.evekit.model.corporation.CustomsOffice;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class CustomsOfficesSheetWriter {

  // Singleton
  private CustomsOfficesSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // CustomsOffices.csv
    // CustomsOfficesMeta.csv
    stream.putNextEntry(new ZipEntry("CustomsOffices.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Item ID", "Solar System ID", "Solar System Name", "Reinforce Hour", "Allow Alliance", "Allow Standings", "Standing Level",
                       "Tax Rate Alliance", "Tax Rate Corp", "Tax Rate Standing High", "Tax Rate Standing Good", "Tax Rate Standing Neutral",
                       "Tax Rate Standing Bad", "Tax Rate Standing Horrible");
    List<Long> metaIDs = new ArrayList<Long>();
    List<CustomsOffice> offices = CustomsOffice.getAll(acct, at);
    for (CustomsOffice next : offices) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getItemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getSolarSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getSolarSystemName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getReinforceHour(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.isAllowAlliance(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.isAllowStandings(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getStandingLevel(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateAlliance(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateCorp(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateStandingHigh(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateStandingGood(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateStandingNeutral(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateStandingBad(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateStandingHorrible(), SheetUtils.CellFormat.DOUBLE_STYLE)); 
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("CustomsOfficesMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CustomsOffice");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
