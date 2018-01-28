package enterprises.orbital.evekit.snapshot.common;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.common.IndustryJob;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class IndustryJobSheetWriter {

  // Singleton
  private IndustryJobSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // IndustryJobs.csv
    // IndustryJobsMeta.csv
    stream.putNextEntry(new ZipEntry("IndustryJobs.csv"));
    final CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Job ID", "Installer ID", "Facility ID", "Station ID", "Activity ID",
                       "Blueprint ID", "Blueprint Type ID", "Blueprint Location ID", "Output Location ID", "Runs",
                       "Cost",
                       "Licensed Runs", "Probability", "Product Type ID", "Status", "Time In Seconds",
                       "Start Date (Raw)", "Start Date",
                       "End Date (Raw)", "End Date", "Pause Date (Raw)", "Pause Date", "Completed Date (Raw)",
                       "Completed Date", "Completed Character ID",
                       "Successful Runs");

    List<Long> metaIDs = new ArrayList<>();
    CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
    CachedData.stream(at, (long contid, AttributeSelector ats) ->
                          IndustryJob.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                                  AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                  AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                  AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                  AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                  AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                  AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                  AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any()),
                      true, capture)
              .forEach(next -> {
                try {
                  //@formatter:off
                  SheetUtils.populateNextRow(output,
                                             new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getJobID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getInstallerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getFacilityID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getStationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getActivityID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getBlueprintID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getBlueprintTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getBlueprintLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getOutputLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getRuns(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getCost(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE),
                                             new DumpCell(next.getLicensedRuns(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getProbability(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                             new DumpCell(next.getProductTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getStatus(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getTimeInSeconds(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getStartDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(new Date(next.getStartDate()), SheetUtils.CellFormat.DATE_STYLE),
                                             new DumpCell(next.getEndDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(new Date(next.getEndDate()), SheetUtils.CellFormat.DATE_STYLE),
                                             new DumpCell(next.getPauseDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(new Date(next.getPauseDate()), SheetUtils.CellFormat.DATE_STYLE),
                                             new DumpCell(next.getCompletedDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(new Date(next.getCompletedDate()), SheetUtils.CellFormat.DATE_STYLE),
                                             new DumpCell(next.getCompletedCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getSuccessfulRuns(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
                  //@formatter:on
                } catch (IOException e) {
                  capture.handle(e);
                }
                metaIDs.add(next.getCid());
              });

    output.flush();
    stream.closeEntry();

    // Handle MetaData
    CSVPrinter metaOutput = SheetUtils.prepForMetaData("IndustryJobsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, metaOutput, next, "IndustryJob");
      if (count > 0) metaOutput.println();
    }
    metaOutput.flush();
    stream.closeEntry();
  }

}
