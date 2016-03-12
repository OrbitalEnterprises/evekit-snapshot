package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.ResearchAgent;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class ResearchAgentSheetWriter {

  // Singleton
  private ResearchAgentSheetWriter() {}

  public static final Comparator<ResearchAgent> ascendingResearchAgentComparator = new Comparator<ResearchAgent>() {

    @Override
    public int compare(
                       ResearchAgent o1,
                       ResearchAgent o2) {
      if (o1.getAgentID() < o2.getAgentID()) return -1;
      if (o1.getAgentID() == o2.getAgentID()) return 0;
      return 1;
    }

  };

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // ResearchAgents.csv
    // ResearchAgentsMeta.csv
    stream.putNextEntry(new ZipEntry("ResearchAgents.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Agent ID", "Current Points", "Points Per Day", "Remainder Points", "Research Start Date (Raw)", "Research Start Date",
                       "Skill Type ID");
    int contid = -1;
    List<ResearchAgent> batch = ResearchAgent.getAllAgents(acct, at, 1000, contid);
    List<Long> metaIDs = new ArrayList<Long>();
    while (batch.size() > 0) {

      for (ResearchAgent next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getAgentID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getCurrentPoints(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                   new DumpCell(next.getPointsPerDay(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                   new DumpCell(next.getRemainderPoints(), SheetUtils.CellFormat.DOUBLE_STYLE), 
                                   new DumpCell(next.getResearchStartDate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(next.getResearchStartDate()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getSkillTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
        // @formatter:on
        metaIDs.add(next.getCid());
      }
      contid = batch.get(batch.size() - 1).getAgentID();
      batch = ResearchAgent.getAllAgents(acct, at, 1000, contid);
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
