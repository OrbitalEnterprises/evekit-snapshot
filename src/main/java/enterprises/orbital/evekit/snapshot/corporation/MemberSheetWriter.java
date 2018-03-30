package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.Member;
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

public class MemberSheetWriter {

  // Singleton
  private MemberSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // Members.csv
    // MembersMeta.csv
    stream.putNextEntry(new ZipEntry("Members.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Character ID");
    List<Long> metaIDs = new ArrayList<>();
    List<Member> batch = CachedData.retrieveAll(at, (contid, at1) -> Member.accessQuery(acct, contid, 1000, false, at1,
                                                                                        AttributeSelector.any()));

    for (Member next : batch) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MembersMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "Member");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
