package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.CharacterMedal;
import enterprises.orbital.evekit.model.character.CharacterMedalGraphic;
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

public class MedalSheetWriter {

  // Singleton
  private MedalSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // Medals.csv
    // MedalsMeta.csv
    // MedalGraphics.csv
    // MedalGraphicsMeta.csv
    stream.putNextEntry(new ZipEntry("Medals.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Medal ID", "Title", "Description", "Corporation ID", "Issued (Raw)", "Issued",
                       "Issuer ID", "Reason", "Status");
    List<Long> metaIDs = new ArrayList<>();
    List<CharacterMedal> medals = CachedData.retrieveAll(at,
                                                         (contid, at1) -> CharacterMedal.accessQuery(acct, contid, 1000,
                                                                                                     false,
                                                                                                     at1,
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any(),
                                                                                                     AttributeSelector.any()));

    for (CharacterMedal next : medals) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getMedalID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTitle(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getDescription(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getIssued(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(new Date(next.getIssued()), SheetUtils.CellFormat.DATE_STYLE), 
                                 new DumpCell(next.getIssuerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getReason(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getStatus(), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MedalsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CharacterMedal");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

    stream.putNextEntry(new ZipEntry("MedalGraphics.csv"));
    output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Medal ID", "Issued (Raw)", "Issued", "Part", "Layer", "Graphic", "Color");
    metaIDs.clear();
    List<CharacterMedalGraphic> medalGraphics = CachedData.retrieveAll(at,
                                                                       (contid, at1) -> CharacterMedalGraphic.accessQuery(
                                                                           acct, contid, 1000, false,
                                                                           at1,
                                                                           AttributeSelector.any(),
                                                                           AttributeSelector.any(),
                                                                           AttributeSelector.any(),
                                                                           AttributeSelector.any(),
                                                                           AttributeSelector.any(),
                                                                           AttributeSelector.any()));

    for (CharacterMedalGraphic next : medalGraphics) {
      // @formatter:off
      SheetUtils.populateNextRow(output,
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getMedalID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getIssued(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getIssued()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getPart(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getLayer(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getGraphic(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getColor(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MedalGraphicsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CharacterMedalGraphic");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

  }

}
