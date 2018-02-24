package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.SkillInQueue;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.w3c.dom.Attr;

public class SkillsInQueueSheetWriter {

  // Singleton
  private SkillsInQueueSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // SkillsInQueue.csv
    // SkillsInQueueMeta.csv
    stream.putNextEntry(new ZipEntry("SkillsInQueue.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Type ID", "Queue Position", "Level", "Start SP", "Start Time (Raw)", "Start Time", "End SP", "End Time (Raw)", "End Time", "Training Start SP");
    List<Long> metaIDs = new ArrayList<>();
    List<SkillInQueue> skill = CachedData.retrieveAll(at, (contid, at1) -> SkillInQueue.accessQuery(acct, contid, 1000, false,
                                                                                                    at1,
                                                                                                    AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                                                                    AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                                                                    AttributeSelector.any(), AttributeSelector.any()));

    for (SkillInQueue next : skill) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getQueuePosition(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getLevel(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getStartSP(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getStartTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(new Date(next.getStartTime()), SheetUtils.CellFormat.DATE_STYLE), 
                                 new DumpCell(next.getEndSP(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getEndTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(new Date(next.getEndTime()), SheetUtils.CellFormat.DATE_STYLE), 
                                 new DumpCell(next.getTrainingStartSP(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("SkillsInQueueMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "SkillInQueue");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
