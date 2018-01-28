package enterprises.orbital.evekit.snapshot.common;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.common.Contract;
import enterprises.orbital.evekit.model.common.ContractBid;
import enterprises.orbital.evekit.model.common.ContractItem;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ContractSheetWriter {

  // Singleton
  private ContractSheetWriter() {}

  private static List<Long> dumpContractItems(SynchronizedEveAccount acct, ZipOutputStream stream,
                                              List<Contract> contracts, long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("ContractItems.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Contract ID", "Record ID", "Type ID", "Quantity", "Raw Quantity", "Singleton",
                       "Included");
    for (Contract nextContract : contracts) {
      int contractID = nextContract.getContractID();
      List<ContractItem> allItems = CachedData.retrieveAll(at, (long contid, AttributeSelector ats) ->
          ContractItem.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.values(contractID),
                                   AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                   AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any()));
      if (!allItems.isEmpty()) {
        for (ContractItem next : allItems) {
          // @formatter:off
          SheetUtils.populateNextRow(output,
                                     new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getContractID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getRecordID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getQuantity(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getRawQuantity(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.isSingleton(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.isIncluded(), SheetUtils.CellFormat.NO_STYLE));
          // @formatter:on
        }
        itemIDs.addAll(allItems.stream()
                               .map(CachedData::getCid)
                               .collect(Collectors.toList()));
        output.println();
      }
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  private static List<Long> dumpContractBids(SynchronizedEveAccount acct, ZipOutputStream stream,
                                             List<Contract> contracts, long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("ContractBids.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Bid ID", "Contract ID", "Bidder ID", "Date Bid (Raw)", "Date Bid", "Amount");
    for (Contract nextContract : contracts) {
      int contractID = nextContract.getContractID();
      List<ContractBid> allBids = CachedData.retrieveAll(at, (long contid, AttributeSelector ats) ->
          ContractBid.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                  AttributeSelector.values(contractID), AttributeSelector.any(),
                                  AttributeSelector.any(), AttributeSelector.any()));
      if (!allBids.isEmpty()) {
        for (ContractBid next : allBids) {
          // @formatter:off
          SheetUtils.populateNextRow(output,
                                     new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getBidID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getContractID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getBidderID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getDateBid(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(new Date(next.getDateBid()), SheetUtils.CellFormat.DATE_STYLE),
                                     new DumpCell(next.getAmount(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE));
          // @formatter:on
        }
        itemIDs.addAll(allBids.stream()
                              .map(CachedData::getCid)
                              .collect(Collectors.toList()));
        output.println();
      }
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  public static void dumpToSheet(SynchronizedEveAccount acct, ZipOutputStream stream, long at) throws IOException {
    // Sections:
    // Contracts.csv
    // ContractsMeta.csv
    // ContractItems.csv
    // ContractItemsMeta.csv
    // ContractBids.csv
    // ContractBidsMeta.csv
    stream.putNextEntry(new ZipEntry("Contracts.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Contract ID", "Issuer ID", "Issuer Corp ID", "Assignee ID", "Acceptor ID",
                       "Start Station ID", "End Station ID", "Type", "Status",
                       "Title", "For Corp", "Availability", "Date Issued (Raw)", "Date Issued", "Date Expired (Raw)",
                       "Date Expired", "Date Accepted (Raw)",
                       "Date Accepted", "Num Days", "Date Completed (Raw)", "Date Completed", "Price", "Reward",
                       "Collateral", "Buyout", "Volume");
    List<Contract> contracts = CachedData.retrieveAll(at, (long contid, AttributeSelector ats) ->
        Contract.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                             AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                             AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                             AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                             AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                             AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                             AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                             AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any()));

    // Write out contract data first
    for (Contract next : contracts) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getContractID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getIssuerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getIssuerCorpID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getAssigneeID(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getAcceptorID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getStartStationID(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getEndStationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getType(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getStatus(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getTitle(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.isForCorp(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getAvailability(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getDateIssued(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getDateIssued()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getDateExpired(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getDateExpired()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getDateAccepted(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getDateAccepted()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getNumDays(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getDateCompleted(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getDateCompleted()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getPrice(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE),
                                 new DumpCell(next.getReward(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE),
                                 new DumpCell(next.getCollateral(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE),
                                 new DumpCell(next.getBuyout(), SheetUtils.CellFormat.BIG_DECIMAL_STYLE),
                                 new DumpCell(next.getVolume(), SheetUtils.CellFormat.DOUBLE_STYLE)); 
      // @formatter:on
    }
    output.flush();
    stream.closeEntry();

    if (contracts.size() > 0) {
      // Wrote at least one contract so proceed

      // Write out meta data for contracts
      output = SheetUtils.prepForMetaData("ContractsMeta.csv", stream, false, null);
      for (Contract next : contracts) {
        int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "Contract");
        if (count > 0) output.println();
      }
      output.flush();
      stream.closeEntry();

      // Write out contract item data in the same style as meta data.
      List<Long> writes = dumpContractItems(acct, stream, contracts, at);
      if (writes.size() > 0) {
        // Only write out meta-data if a contract Item was actually written.
        output = SheetUtils.prepForMetaData("ContractItemsMeta.csv", stream, false, null);
        for (Long next : writes) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "ContractItem");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }

      // Write out contract bid data in the same style as meta data.
      writes = dumpContractBids(acct, stream, contracts, at);
      if (writes.size() > 0) {
        // Only write out meta-data if a contract Item was actually written.
        output = SheetUtils.prepForMetaData("ContractBidsMeta.csv", stream, false, null);
        for (Long next : writes) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "ContractBid");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }
    }
  }

}
