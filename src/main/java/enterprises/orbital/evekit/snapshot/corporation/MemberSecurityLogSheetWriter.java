package enterprises.orbital.evekit.snapshot.corporation;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.corporation.MemberSecurityLog;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class MemberSecurityLogSheetWriter {

  // Singleton
  private MemberSecurityLogSheetWriter() {}

  public static String setToString(
                                   Set<Long> convert) {
    return Arrays.toString(convert.toArray(new Long[convert.size()]));
  }

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // MemberSecurityLogs.csv
    // MemberSecurityLogsMeta.csv
    stream.putNextEntry(new ZipEntry("MemberSecurityLogs.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Change Time (Raw)", "Change Time", "Changed Character ID", "Changed Character Name", "Issuer ID", "Issuer Name",
                       "Role Location Type", "Old Roles", "New Roles");
    List<Long> metaIDs = new ArrayList<Long>();
    long contid = -1;
    List<MemberSecurityLog> batch = MemberSecurityLog.getAllForward(acct, at, 1000, contid);

    while (batch.size() > 0) {

      for (MemberSecurityLog next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getChangeTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(next.getChangeTime()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getChangedCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getChangedCharacterName(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getIssuerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getIssuerName(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getRoleLocationType(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(setToString(next.getOldRoles()), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(setToString(next.getNewRoles()), SheetUtils.CellFormat.NO_STYLE)); 
        // @formatter:on
        metaIDs.add(next.getCid());
      }

      contid = batch.get(batch.size() - 1).getChangeTime();
      batch = MemberSecurityLog.getAllForward(acct, at, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MemberSecurityLogsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "MemberSecurityLog");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
