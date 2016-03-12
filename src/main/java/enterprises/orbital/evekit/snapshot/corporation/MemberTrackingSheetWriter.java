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
import enterprises.orbital.evekit.model.corporation.MemberTracking;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class MemberTrackingSheetWriter {

  // Singleton
  private MemberTrackingSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // MemberTracking.csv
    // MemberTrackingMeta.csv
    stream.putNextEntry(new ZipEntry("MemberTracking.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Character ID", "Base", "Base ID", "Grantable Roles", "Location", "Location ID", "Logoff Date Time (Raw)", "Logoff Date Time",
                       "Logon Date Time (Raw)", "Logon Date Time", "Name", "Roles", "Ship Type", "Ship Type ID", "Start Date Time (Raw)", "Start Date Time",
                       "Title");
    List<Long> metaIDs = new ArrayList<Long>();
    long contid = -1;
    List<MemberTracking> batch = MemberTracking.getAll(acct, at, 1000, contid);

    while (batch.size() > 0) {

      for (MemberTracking next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getBase(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getBaseID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getGrantableRoles(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getLocation(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getLogoffDateTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getLogoffDateTime()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(next.getLogonDateTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getLogonDateTime()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(next.getName(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getRoles(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getShipType(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getShipTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getStartDateTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getStartDateTime()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(next.getTitle(), SheetUtils.CellFormat.NO_STYLE)); 
        // @formatter:on
        metaIDs.add(next.getCid());
      }

      contid = batch.get(batch.size() - 1).getCharacterID();
      batch = MemberTracking.getAll(acct, at, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MemberTrackingMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "MemberTracking");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
