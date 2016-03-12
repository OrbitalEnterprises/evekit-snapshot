package enterprises.orbital.evekit.snapshot.common;

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
import enterprises.orbital.evekit.model.common.IndustryJob;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

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
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Job ID", "Installer ID", "Installer Name", "Facility ID", "Solar System ID", "Solar System Name", "Station ID", "Activity ID",
                       "Blueprint ID", "Blueprint Type ID", "Blueprint Type Name", "Blueprint Location ID", "Output Location ID", "Runs", "Cost", "Team ID",
                       "Licensed Runs", "Probability", "Product Type ID", "Product Type Name", "Status", "Time In Seconds", "Start Date (Raw)", "Start Date",
                       "End Date (Raw)", "End Date", "Pause Date (Raw)", "Pause Date", "Completed Date (Raw)", "Completed Date", "Completed Character ID",
                       "Successful Runs");

    long contid = -1;
    List<Long> metaIDs = new ArrayList<Long>();
    List<IndustryJob> batch = IndustryJob.getAllForward(acct, at, 1000, contid);

    while (batch.size() > 0) {

      for (IndustryJob next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getJobID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getInstallerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getInstallerName(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getFacilityID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getSolarSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getSolarSystemName(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getStationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getActivityID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getBlueprintID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getBlueprintTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getBlueprintTypeName(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getBlueprintLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getOutputLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getRuns(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getCost(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE), 
                                   new DumpCell(next.getTeamID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getLicensedRuns(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getProbability(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                   new DumpCell(next.getProductTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getProductTypeName(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getStatus(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
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
        // @formatter:on
        metaIDs.add(next.getCid());
      }

      contid = batch.get(batch.size() - 1).getStartDate();
      batch = IndustryJob.getAllForward(acct, at, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("IndustryJobsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "IndustryJob");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
