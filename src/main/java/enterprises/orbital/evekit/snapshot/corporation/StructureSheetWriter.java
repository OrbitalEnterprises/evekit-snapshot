package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.Structure;
import enterprises.orbital.evekit.model.corporation.StructureService;
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

public class StructureSheetWriter {

  // Singleton
  private StructureSheetWriter() {}

  private static List<Long> dumpServices(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("StructureServices.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Structure ID", "Name", "State");

    List<StructureService> services = CachedData.retrieveAll(at,
                                                             (contid, at1) -> StructureService.accessQuery(acct, contid,
                                                                                                           1000, false,
                                                                                                           at1,
                                                                                                           AttributeSelector.any(),
                                                                                                           AttributeSelector.any(),
                                                                                                           AttributeSelector.any()));
    for (StructureService next : services) {
      // @formatter:off
          SheetUtils.populateNextRow(output,
                                     new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getStructureID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getName(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getState(), SheetUtils.CellFormat.NO_STYLE));
          // @formatter:on
      itemIDs.add(next.getCid());
    }

    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // Structures.csv
    // StructuresMeta.csv
    // StructureServices.csv
    // StructureServicesMeta.csv
    stream.putNextEntry(new ZipEntry("Structures.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Structure ID", "Corporation ID", "Fuel Expires (Raw)", "Fuel Expires",
                       "Next Reinforce Apply (Raw)", "Next Reinforce Apply", "Next Reinforce Hour",
                       "Next Reinforce Weekday", "Profile ID", "Reinforce Hour", "Reinforce Weekday",
                       "State", "State Timer End (Raw)", "State Timer End", "State Timer Start (Raw)",
                       "State Timer Start", "System ID", "Type ID", "Unanchors At (Raw)",
                       "Unanchors At");
    List<Structure> structures = CachedData.retrieveAll(at,
                                                        (contid, at1) -> Structure.accessQuery(acct, contid, 1000,
                                                                                               false, at1,
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any(),
                                                                                               AttributeSelector.any()));


    for (Structure next : structures) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getStructureID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getFuelExpires(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getFuelExpires()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getNextReinforceApply(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getNextReinforceApply()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getNextReinforceHour(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getNextReinforceWeekday(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getProfileID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getReinforceHour(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getReinforceWeekday(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getState(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getStateTimerEnd(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getStateTimerEnd()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getStateTimerStart(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getStateTimerStart()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getUnanchorsAt(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getUnanchorsAt()), SheetUtils.CellFormat.DATE_STYLE));
    }
    output.flush();
    stream.closeEntry();

    if (structures.size() > 0) {
      // Write out meta data for structures
      output = SheetUtils.prepForMetaData("StructuresMeta.csv", stream, false, null);
      for (Structure next : structures) {
        int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "Structure");
        if (count > 0) output.println();
      }
      output.flush();
      stream.closeEntry();

      // Write out structure services data in the same style as meta data.
      List<Long> writes = dumpServices(acct, stream, at);
      if (writes.size() > 0) {
        // Only write out meta-data if service was actually written.
        output = SheetUtils.prepForMetaData("StructureServicesMeta.csv", stream, false, null);
        for (Long next : writes) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "StructureService");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }
    }
  }

}
