package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.ResearchAgent;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ResearchAgentSheetWriter {

  // Singleton
  private ResearchAgentSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // ResearchAgents.csv
    // ResearchAgentsMeta.csv
    stream.putNextEntry(new ZipEntry("ResearchAgents.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Agent ID", "Points Per Day", "Remainder Points", "Research Start Date (Raw)", "Research Start Date",
                       "Skill Type ID");
    List<ResearchAgent> batch = SheetUtils.retrieveAll(at, (ctid, tm) -> ResearchAgent.accessQuery(acct, ctid, 1000, false, tm, SheetUtils.ANY_SELECTOR, SheetUtils.ANY_SELECTOR, SheetUtils.ANY_SELECTOR, SheetUtils.ANY_SELECTOR, SheetUtils.ANY_SELECTOR));
    List<Long> metaIDs = new ArrayList<>();
    for (ResearchAgent next : batch) {
      // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getAgentID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getPointsPerDay(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                   new DumpCell(next.getRemainderPoints(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                   new DumpCell(next.getResearchStartDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(next.getResearchStartDate()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getSkillTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
        // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("ResearchAgentsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "ResearchAgent");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
