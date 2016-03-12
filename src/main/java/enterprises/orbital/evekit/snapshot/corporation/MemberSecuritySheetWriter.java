package enterprises.orbital.evekit.snapshot.corporation;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.corporation.MemberSecurity;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class MemberSecuritySheetWriter {

  // Singleton
  private MemberSecuritySheetWriter() {}

  public static String setToString(
                                   Set<Long> convert) {
    return Arrays.toString(convert.toArray(new Long[convert.size()]));
  }

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // MemberSecurity.csv
    // MemberSecurityMeta.csv
    stream.putNextEntry(new ZipEntry("MemberSecurity.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Character ID", "Name", "Grantable Roles", "Grantable Roles At Base", "Grantable Roles At HQ", "Grantable Roles At Other", "Roles",
                       "Roles At Base", "Roles At HQ", "Roles At Other", "Titles");
    List<Long> metaIDs = new ArrayList<Long>();
    long contid = -1;
    List<MemberSecurity> batch = MemberSecurity.getAll(acct, at, 1000, contid);

    while (batch.size() > 0) {

      for (MemberSecurity next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getName(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(setToString(next.getGrantableRoles()), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(setToString(next.getGrantableRolesAtBase()), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(setToString(next.getGrantableRolesAtHQ()), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(setToString(next.getGrantableRolesAtOther()), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(setToString(next.getRoles()), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(setToString(next.getRolesAtBase()), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(setToString(next.getRolesAtHQ()), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(setToString(next.getRolesAtOther()), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(setToString(next.getTitles()), SheetUtils.CellFormat.NO_STYLE)); 
        // @formatter:on
        metaIDs.add(next.getCid());
      }

      contid = batch.get(batch.size() - 1).getCharacterID();
      batch = MemberSecurity.getAll(acct, at, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MemberSecurityMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "MemberSecurity");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
