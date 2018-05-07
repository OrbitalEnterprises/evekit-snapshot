package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.MiningExtraction;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class MiningExtractionSheetWriter {

  // Singleton
  private MiningExtractionSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // MiningExtraction.csv
    // MiningExtractionMeta.csv
    stream.putNextEntry(new ZipEntry("MiningExtraction.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Moon ID", "Structure ID", "Extraction Start Time (Raw)",
                       "Extraction Start Time", "Chunk Arrival Time (Raw)", "Chunk Arrival Time",
                       "Natural Decay Time (Raw)", "Natural Decay Time");
    List<MiningExtraction> points = CachedData.retrieveAll(at,
                                                           (contid, at1) -> MiningExtraction.accessQuery(acct, contid,
                                                                                                         1000,
                                                                                                         false,
                                                                                                         at1,
                                                                                                         AttributeSelector.any(),
                                                                                                         AttributeSelector.any(),
                                                                                                         AttributeSelector.any(),
                                                                                                         AttributeSelector.any(),
                                                                                                         AttributeSelector.any()));

    for (MiningExtraction next : points) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getMoonID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getStructureID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getExtractionStartTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getExtractionStartTime()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getChunkArrivalTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getChunkArrivalTime()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getNaturalDecayTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getNaturalDecayTime()), SheetUtils.CellFormat.DATE_STYLE));
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MiningExtractionMeta.csv", stream, false, null);
    for (MiningExtraction next : points) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "MiningExtraction");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

  }

}
