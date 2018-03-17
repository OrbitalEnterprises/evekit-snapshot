package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.CharacterTitle;
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

public class TitleSheetWriter {

  // Singleton
  private TitleSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // Titles.csv
    // TitlesMeta.csv
    stream.putNextEntry(new ZipEntry("Titles.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Title ID", "Title Name");
    List<Long> metaIDs = new ArrayList<>();
    List<CharacterTitle> titles = CachedData.retrieveAll(at,
                                                         (contid, at1) -> CharacterTitle.accessQuery(acct, contid, 1000,
                                                                                                     false,
                                                                                                     at1,
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any()));
    for (CharacterTitle next : titles) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getTitleID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTitleName(), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("TitlesMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CharacterTitle");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
