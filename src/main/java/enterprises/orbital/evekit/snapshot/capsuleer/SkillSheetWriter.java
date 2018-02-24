package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.CharacterSkill;
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
    final CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Type ID", "Trained Skill Level", "Skill Points", "Active Skill Level");

    List<Long> metaIDs = new ArrayList<>();
    CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
    CachedData.stream(at, (long contid, AttributeSelector ats) ->
                          CharacterSkill.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                               AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any()),
                      true, capture)
              .forEach(next -> {
                try {
                  // @formatter:off
                  SheetUtils.populateNextRow(output,
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getTrainedSkillLevel(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getSkillpoints(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getActiveSkillLevel(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
                  // @formatter:on
                } catch (IOException e) {
                  capture.handle(e);
                }
                metaIDs.add(next.getCid());
              });


    // Handle MetaData
    CSVPrinter outputMeta = SheetUtils.prepForMetaData("SkillsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, outputMeta, next, "CharacterSkill");
      if (count > 0) output.println();
    }
    outputMeta.flush();
    stream.closeEntry();
  }

}
