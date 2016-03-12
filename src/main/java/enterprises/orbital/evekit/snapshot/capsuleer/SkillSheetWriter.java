package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterSkill;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class SkillSheetWriter {

  // Singleton
  private SkillSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Skills.csv
    // SkillsMeta.csv
    stream.putNextEntry(new ZipEntry("Skills.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Type ID", "Level", "Skill Points", "Published");
    List<Long> metaIDs = new ArrayList<Long>();
    int contid = -1;
    List<CharacterSkill> batch = CharacterSkill.getAll(acct, at, 1000, contid);
    while (batch.size() > 0) {

      for (CharacterSkill next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getLevel(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getSkillpoints(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.isPublished(), SheetUtils.CellFormat.NO_STYLE)); 
        // @formatter:on
        metaIDs.add(next.getCid());
      }
      contid = batch.get(batch.size() - 1).getTypeID();
      batch = CharacterSkill.getAll(acct, at, 1000, contid);
    }

    // Handle MetaData
    output = SheetUtils.prepForMetaData("SkillsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CharacterSkill");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
