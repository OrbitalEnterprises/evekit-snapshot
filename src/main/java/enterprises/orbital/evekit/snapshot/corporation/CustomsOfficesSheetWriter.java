package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.CustomsOffice;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

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
    output.printRecord("ID", "Office ID", "Solar System ID", "Reinforce Exit Start", "Reinforce Exit End",
                       "Allow Alliance", "Allow Standings", "Standing Level",
                       "Tax Rate Alliance", "Tax Rate Corp", "Tax Rate Standing Excellent", "Tax Rate Standing Good",
                       "Tax Rate Standing Neutral",
                       "Tax Rate Standing Bad", "Tax Rate Standing Terrible");
    List<Long> metaIDs = new ArrayList<>();
    List<CustomsOffice> offices = CachedData.retrieveAll(at,
                                                         (contid, at1) -> CustomsOffice.accessQuery(acct, contid, 1000,
                                                                                                    false,
                                                                                                    at1,
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any()));
    for (CustomsOffice next : offices) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getOfficeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getSolarSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getReinforceExitStart(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getReinforceExitEnd(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.isAllowAlliance(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.isAllowStandings(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getStandingLevel(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getTaxRateAlliance(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateCorp(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateStandingExcellent(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                 new DumpCell(next.getTaxRateStandingGood(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateStandingNeutral(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateStandingBad(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                 new DumpCell(next.getTaxRateStandingTerrible(), SheetUtils.CellFormat.DOUBLE_STYLE));
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
