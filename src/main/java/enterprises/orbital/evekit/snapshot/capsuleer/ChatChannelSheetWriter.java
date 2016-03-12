package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.ChatChannel;
import enterprises.orbital.evekit.model.character.ChatChannelMember;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class ChatChannelSheetWriter {

  // Singleton
  private ChatChannelSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // ChatChannels.csv
    // ChatChannelsMeta.csv
    // ChatChannelMembers.csv
    // ChatChannelMembersMeta.csv
    stream.putNextEntry(new ZipEntry("ChatChannels.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Channel ID", "Owner ID", "Owner Name", "Display Name", "Comparison Key", "Has Password", "MOTD");
    List<ChatChannel> channels = ChatChannel.getAllChatChannels(acct, at);

    for (ChatChannel next : channels) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getChannelID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getOwnerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getOwnerName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getDisplayName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getComparisonKey(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.isHasPassword(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getMotd(), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("ChatChannelsMeta.csv", stream, false, null);
    for (ChatChannel next : channels) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "ChatChannel");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

    stream.putNextEntry(new ZipEntry("ChatChannelMembers.csv"));
    output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Channel ID", "Category", "Accessor ID", "Accessor Name", "Until When (Raw)", "Until When", "Reason");
    List<ChatChannelMember> members = ChatChannelMember.getAllChatChannelMembers(acct, at);

    for (ChatChannelMember next : members) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getChannelID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getCategory(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getAccessorID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getAccessorName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getUntilWhen(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getUntilWhen()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getReason(), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
    }

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("ChatChannelMembersMeta.csv", stream, false, null);
    for (ChatChannelMember next : members) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "ChatChannelMember");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

  }

}
