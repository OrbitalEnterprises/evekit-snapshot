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
import enterprises.orbital.evekit.model.corporation.CorporationTitle;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class CorporationTitleSheetWriter {

  // Singleton
  private CorporationTitleSheetWriter() {}

  public static String setToString(
                                   Set<Long> convert) {
    return Arrays.toString(convert.toArray(new Long[convert.size()]));
  }

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Titles.csv
    // TitlesMeta.csv
    stream.putNextEntry(new ZipEntry("Titles.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Title ID", "Title Name", "Grantable Roles", "Grantable Roles At Base", "Grantable Roles At HQ", "Grantable Roles At Other",
                       "Roles", "Roles At Base", "Roles At HQ", "Roles At Other");
    List<Long> metaIDs = new ArrayList<Long>();
    List<CorporationTitle> batch = CorporationTitle.getAll(acct, at);

    for (CorporationTitle next : batch) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getTitleID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTitleName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(setToString(next.getGrantableRoles()), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(setToString(next.getGrantableRolesAtBase()), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(setToString(next.getGrantableRolesAtHQ()), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(setToString(next.getGrantableRolesAtOther()), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(setToString(next.getRoles()), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(setToString(next.getRolesAtBase()), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(setToString(next.getRolesAtHQ()), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(setToString(next.getRolesAtOther()), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("TitlesMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CorporationTitle");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
