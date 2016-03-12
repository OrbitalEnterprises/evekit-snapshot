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
import enterprises.orbital.evekit.model.corporation.Role;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class CorporationRoleSheetWriter {

  // Singleton
  private CorporationRoleSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Roles.csv
    // RolesMeta.csv
    stream.putNextEntry(new ZipEntry("Roles.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Role ID", "Role Description", "Role Name");
    List<Long> metaIDs = new ArrayList<Long>();
    List<Role> roles = Role.getAll(acct, at);
    for (Role next : roles) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getRoleID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getRoleDescription(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getRoleName(), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("RolesMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "Role");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
