package enterprises.orbital.evekit.snapshot.common;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.common.Contract;
import enterprises.orbital.evekit.model.common.ContractBid;
import enterprises.orbital.evekit.model.common.ContractItem;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class ContractSheetWriter {

  // Singleton
  private ContractSheetWriter() {}

  public static List<Long> dumpContractItems(
                                             SynchronizedEveAccount acct,
                                             ZipOutputStream stream,
                                             List<Contract> contracts,
                                             long at)
    throws IOException {
    List<Long> itemIDs = new ArrayList<Long>();
    stream.putNextEntry(new ZipEntry("ContractItems.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Contract ID", "Record ID", "Type ID", "Quantity", "Raw Quantity", "Singleton", "Included");
    for (Contract nextContract : contracts) {
      long contractID = nextContract.getContractID();
      List<ContractItem> allItems = new ArrayList<ContractItem>();
      long contid = -1;
      List<ContractItem> batch = ContractItem.getAllContractItems(acct, at, contractID, 1000, contid);
      while (batch.size() > 0) {
        allItems.addAll(batch);
        contid = batch.get(batch.size() - 1).getRecordID();
        batch = ContractItem.getAllContractItems(acct, at, contractID, 1000, contid);
      }
      if (allItems.size() > 0) {
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
          itemIDs.add(next.getCid());
        }
        output.println();
      }
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  public static List<Long> dumpContractBids(
                                            SynchronizedEveAccount acct,
                                            ZipOutputStream stream,
                                            List<Contract> contracts,
                                            long at)
    throws IOException {
    List<Long> itemIDs = new ArrayList<Long>();
    stream.putNextEntry(new ZipEntry("ContractBids.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Bid ID", "Contract ID", "Bidder ID", "Date Bid (Raw)", "Date Bid", "Amount");
    for (Contract nextContract : contracts) {
      long contractID = nextContract.getContractID();
      List<ContractBid> allBids = new ArrayList<ContractBid>();
      long contid = -1;
      List<ContractBid> batch = ContractBid.getAllBidsByContractID(acct, at, contractID, 1000, contid);
      while (batch.size() > 0) {
        allBids.addAll(batch);
        contid = batch.get(batch.size() - 1).getBidID();
        batch = ContractBid.getAllBidsByContractID(acct, at, contractID, 1000, contid);
      }
      if (allBids.size() > 0) {
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
          itemIDs.add(next.getCid());
        }
        output.println();
      }
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at)
    throws IOException {
    // Sections:
    // Contracts.csv
    // ContractsMeta.csv
    // ContractItems.csv
    // ContractItemsMeta.csv
    // ContractBids.csv
    // ContractBidsMeta.csv
    stream.putNextEntry(new ZipEntry("Contracts.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Contract ID", "Issuer ID", "Issuer Corp ID", "Assignee ID", "Acceptor ID", "Start Station ID", "End Station ID", "Type", "Status",
                       "Title", "For Corp", "Availability", "Date Issued (Raw)", "Date Issued", "Date Expired (Raw)", "Date Expired", "Date Accepted (Raw)",
                       "Date Accepted", "Num Days", "Date Completed (Raw)", "Date Completed", "Price", "Reward", "Collateral", "Buyout", "Volume");
    List<Contract> contracts = new ArrayList<Contract>();
    long contid = 0;
    List<Contract> batch = Contract.getAllContracts(acct, at, 1000, contid);
    while (batch.size() > 0) {
      contracts.addAll(batch);
      contid = batch.get(batch.size() - 1).getContractID();
      batch = Contract.getAllContracts(acct, at, 1000, contid);
    }

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
