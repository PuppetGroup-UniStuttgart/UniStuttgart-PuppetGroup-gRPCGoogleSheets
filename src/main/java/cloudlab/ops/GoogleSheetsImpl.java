package cloudlab.ops;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.List;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.PlainTextConstruct;
import com.google.gdata.data.spreadsheet.CellEntry;
import com.google.gdata.data.spreadsheet.CellFeed;
import com.google.gdata.data.spreadsheet.ListEntry;
import com.google.gdata.data.spreadsheet.ListFeed;
import com.google.gdata.data.spreadsheet.SpreadsheetEntry;
import com.google.gdata.data.spreadsheet.SpreadsheetFeed;
import com.google.gdata.data.spreadsheet.WorksheetEntry;
import com.google.gdata.data.spreadsheet.WorksheetFeed;
import com.google.gdata.util.ServiceException;

import cloudlab.GoogleSheetsProto.CreateWorksheetReply;
import cloudlab.GoogleSheetsProto.CreateWorksheetRequest;
import cloudlab.GoogleSheetsProto.DeleteRowReply;
import cloudlab.GoogleSheetsProto.DeleteRowRequest;
import cloudlab.GoogleSheetsProto.DeleteWorksheetReply;
import cloudlab.GoogleSheetsProto.DeleteWorksheetRequest;
import cloudlab.GoogleSheetsProto.Empty;
import cloudlab.GoogleSheetsProto.FetchRowColReply;
import cloudlab.GoogleSheetsProto.FetchRowColRequest;
import cloudlab.GoogleSheetsProto.GetWorksheetContentsReply;
import cloudlab.GoogleSheetsProto.GetWorksheetContentsRequest;
import cloudlab.GoogleSheetsProto.GoogleSheetsGrpc.GoogleSheets;
import cloudlab.GoogleSheetsProto.SpreadsheetListReply;
import cloudlab.GoogleSheetsProto.SpreadsheetListReply.Builder;
import cloudlab.GoogleSheetsProto.UpdateWorksheetReply;
import cloudlab.GoogleSheetsProto.UpdateWorksheetRequest;
import cloudlab.GoogleSheetsProto.WorksheetListReply;
import cloudlab.GoogleSheetsProto.WorksheetListRequest;
import cloudlab.GoogleSheetsProto.*;
import io.grpc.stub.StreamObserver;

public class GoogleSheetsImpl implements GoogleSheets {

	// The name of the client ID of the grpc application service account
	private final String CLIENT_ID = "grpcgooglesheets@ivory-plane-135618.iam.gserviceaccount.com";

	// The name of the p12 file of the grpc application service account
	private final String P12FILE = "/GrpcGoogleSheets.p12";

	private GoogleCredential getCredentials() {

		GoogleCredential credential = null;

		try {
			JacksonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
			HttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();

			String[] SCOPESArray = { "https://spreadsheets.google.com/feeds",
					"https://spreadsheets.google.com/feeds/spreadsheets/private/full",
					"https://docs.google.com/feeds" };
			final List SCOPES = Arrays.asList(SCOPESArray);

			credential = new GoogleCredential.Builder().setTransport(httpTransport).setJsonFactory(JSON_FACTORY)
					.setServiceAccountId(CLIENT_ID).setServiceAccountPrivateKeyFromP12File(new File(P12FILE))
					.setServiceAccountScopes(SCOPES).build();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (GeneralSecurityException e) {
			e.printStackTrace();
		} 
		return credential;
	}

	@SuppressWarnings("null")
	@Override
	public void spreadsheetList(Empty request, StreamObserver<SpreadsheetListReply> responseObserver) {
		SpreadsheetService service = new SpreadsheetService("GoogleSheetsGrpc");
		Builder builder = SpreadsheetListReply.newBuilder();

		try {
			GoogleCredential credential = getCredentials();
			service.setOAuth2Credentials(credential);

			// Define the URL to request. This should never change.
			URL SPREADSHEET_FEED_URL = new URL("https://spreadsheets.google.com/feeds/spreadsheets/private/full");

			// Make a request to the API and get all spreadsheets.
			SpreadsheetFeed feed = service.getFeed(SPREADSHEET_FEED_URL, SpreadsheetFeed.class);
			List<SpreadsheetEntry> spreadsheets = feed.getEntries();

			if (spreadsheets.size() == 0) {
				builder = builder.addSpreadsheetName("No spreadsheets found.");
			}

			// Iterate through all of the spreadsheets returned
			for (SpreadsheetEntry spreadsheet : spreadsheets) {
				// Send the list of spreadsheet names
				builder = builder.addSpreadsheetName(spreadsheet.getTitle().getPlainText().toString());
			}
			
			responseObserver.onNext(builder.build());
			responseObserver.onCompleted();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (ServiceException e) {
			e.printStackTrace();
		}
	}

	@Override
	public void worksheetList(WorksheetListRequest request, StreamObserver<WorksheetListReply> responseObserver) {

		String url = "https://spreadsheets.google.com/feeds/worksheets/" + request.getSpreadsheetId() + "/private/full";

		SpreadsheetService service = new SpreadsheetService("GoogleSheetsGrpc");
		WorksheetListReply.Builder builder = WorksheetListReply.newBuilder();

		String details = null;

		try {
			GoogleCredential credential = getCredentials();
			service.setOAuth2Credentials(credential);

			URL WORKSHEET_FEED_URL = new URL(url);
			WorksheetFeed feed = service.getFeed(WORKSHEET_FEED_URL, WorksheetFeed.class);

			// Make a request to the API to fetch information about all
			// worksheets in the spreadsheet.
			List<WorksheetEntry> worksheets = feed.getEntries();

			// Iterate through each worksheet in the spreadsheet.
			for (WorksheetEntry worksheet : worksheets) {
				// Get the worksheet's title, row count, and column count and
				// send the details.
				details = "\t" + worksheet.getTitle().getPlainText() + "- rows:" + worksheet.getRowCount() + " - cols: "
						+ worksheet.getColCount() + "\n";
				builder = builder.addWorksheetDetails(details);
			}
			responseObserver.onNext(builder.build());
			responseObserver.onCompleted();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (ServiceException e) {
			e.printStackTrace();
		}
	}

	@Override
	public void createWorksheet(CreateWorksheetRequest request, StreamObserver<CreateWorksheetReply> responseObserver) {

		SpreadsheetService service = new SpreadsheetService("GoogleSheetsGrpc");
		CreateWorksheetReply reply = null;

		try {
			GoogleCredential credential = getCredentials();
			service.setOAuth2Credentials(credential);

			// Define the URL to request.
			URL SPREADSHEET_FEED_URL = new URL("https://spreadsheets.google.com/feeds/spreadsheets/private/full");

			// Make a request to the API and get all spreadsheets.
			SpreadsheetFeed feed = service.getFeed(SPREADSHEET_FEED_URL, SpreadsheetFeed.class);
			List<SpreadsheetEntry> spreadsheets = feed.getEntries();

			if (spreadsheets.size() == 0) {
				reply.newBuilder().setSuccess("No spreadsheets found.").build();
			}

			SpreadsheetEntry spreadsheet = null;
			for (int i = 0; i < spreadsheets.size(); i++) {
				spreadsheet = spreadsheets.get(i);
				// Get the required spreadsheet.
				if (spreadsheet.getTitle().getPlainText().equals(request.getSpreadsheetName()))
					break;
			}

			// Create a local representation of the new worksheet.
			WorksheetEntry worksheet = new WorksheetEntry();
			worksheet.setTitle(new PlainTextConstruct(request.getWorksheetName()));
			worksheet.setColCount(request.getRow());
			worksheet.setRowCount(request.getCol());

			// Send the local representation of the worksheet to the API for
			// creation. The URL to use here is the worksheet feed URL of the
			// spreadsheet.
			URL worksheetFeedUrl = spreadsheet.getWorksheetFeedUrl();
			WorksheetEntry worksheetSucess = new WorksheetEntry();
			worksheetSucess = service.insert(worksheetFeedUrl, worksheet);

			if (worksheetSucess.getId() != null)
				reply.newBuilder()
						.setSuccess("Worksheet " + worksheetSucess.getTitle().getPlainText() + " is created. Please check")
						.build();
			else
				reply.newBuilder().setSuccess("Creation unsuccessful. ").build();

			responseObserver.onNext(reply);
			responseObserver.onCompleted();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (ServiceException e) {
			e.printStackTrace();
		}

	}

	@Override
	public void updateWorksheet(UpdateWorksheetRequest request, StreamObserver<UpdateWorksheetReply> responseObserver) {

		String url = "https://spreadsheets.google.com/feeds/worksheets/" + request.getSpreadsheetId() + "/private/full";

		SpreadsheetService service = new SpreadsheetService("GoogleSheetsGrpc");
		UpdateWorksheetReply reply = null;

		try {
			GoogleCredential credential = getCredentials();
			service.setOAuth2Credentials(credential);

			URL WORKSHEET_FEED_URL = new URL(url);
			WorksheetFeed feed = service.getFeed(WORKSHEET_FEED_URL, WorksheetFeed.class);

			// Make a request to the API to fetch information about all
			// worksheets in the spreadsheet.
			List<WorksheetEntry> worksheets = feed.getEntries();

			WorksheetEntry worksheet = null;
			for (int i = 0; i < worksheets.size(); i++) {
				worksheet = worksheets.get(i);
				// Get the required worksheet.
				if (worksheet.getTitle().getPlainText().equals(request.getOldWorksheetName()))
					break;
			}

			// Update the local representation of the worksheet.
			worksheet.setTitle(new PlainTextConstruct(request.getNewWorksheetName()));
			worksheet.setColCount(request.getNewWorksheetRow());
			worksheet.setRowCount(request.getNewWorksheetCol());

			// Send the local representation of the worksheet to the API for
			// modification.
			worksheet.update();
			reply.newBuilder().setSuccess("Worksheet " + worksheet.getTitle().getPlainText() + " is updated. Please check")
					.build();

			responseObserver.onNext(reply);
			responseObserver.onCompleted();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (ServiceException e) {
			e.printStackTrace();
		}
	}

	@Override
	public void deleteWorksheet(DeleteWorksheetRequest request, StreamObserver<DeleteWorksheetReply> responseObserver) {
		String url = "https://spreadsheets.google.com/feeds/worksheets/" + request.getSpreadsheetId() + "/private/full";

		SpreadsheetService service = new SpreadsheetService("GoogleSheetsGrpc");
		DeleteWorksheetReply reply = null;

		try {
			GoogleCredential credential = getCredentials();
			service.setOAuth2Credentials(credential);

			URL WORKSHEET_FEED_URL = new URL(url);
			WorksheetFeed feed = service.getFeed(WORKSHEET_FEED_URL, WorksheetFeed.class);

			// Make a request to the API to fetch information about all
			// worksheets in the spreadsheet.
			List<WorksheetEntry> worksheets = feed.getEntries();

			WorksheetEntry worksheet = null;
			for (int i = 0; i < worksheets.size(); i++) {
				worksheet = worksheets.get(i);
				// Get the required worksheet.
				if (worksheet.getTitle().getPlainText().equals(request.getWorksheetName()))
					break;
			}

			// Delete the worksheet via the API.
			worksheet.delete();
			reply.newBuilder().setSuccess("Worksheet " + worksheet.getTitle().getPlainText() + " is deleted. Please check")
					.build();

			responseObserver.onNext(reply);
			responseObserver.onCompleted();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (ServiceException e) {
			e.printStackTrace();
		}

	}

	@Override
	public void getWorksheetContents(GetWorksheetContentsRequest request,
			StreamObserver<GetWorksheetContentsReply> responseObserver) {
		String url = "https://spreadsheets.google.com/feeds/worksheets/" + request.getSpreadsheetId() + "/private/full";

		SpreadsheetService service = new SpreadsheetService("GoogleSheetsGrpc");
		GetWorksheetContentsReply.Builder builder = GetWorksheetContentsReply.newBuilder();

		String details = null;

		try {
			GoogleCredential credential = getCredentials();
			service.setOAuth2Credentials(credential);

			URL WORKSHEET_FEED_URL = new URL(url);
			WorksheetFeed feed = service.getFeed(WORKSHEET_FEED_URL, WorksheetFeed.class);

			// Make a request to the API to fetch information about all
			// worksheets in the spreadsheet.
			List<WorksheetEntry> worksheets = feed.getEntries();

			WorksheetEntry worksheet = null;
			for (int i = 0; i < worksheets.size(); i++) {
				worksheet = worksheets.get(i);
				// Get the required worksheet.
				if (worksheet.getTitle().getPlainText().equals(request.getWorksheetName()))
					break;
			}

			URL cellFeedUrl = new URL(worksheet.getCellFeedUrl().toString() + "?min-row=1&max-row=1");
			CellFeed cellFeed = service.getFeed(cellFeedUrl, CellFeed.class);
			for (CellEntry cellEntry : cellFeed.getEntries()) {
				details = cellEntry.getCell().getValue() + "\t";
			}
			builder = builder.addContents(details + "\n");

			// Fetch the list feed of the worksheet.
			URL listFeedUrl = worksheet.getListFeedUrl();
			ListFeed listFeed = service.getFeed(listFeedUrl, ListFeed.class);

			// Iterate through each row, printing its cell values.
			for (ListEntry row : listFeed.getEntries()) {
				// Iterate over the remaining columns, and print each cell value
				for (String tag : row.getCustomElements().getTags()) {
					details = row.getCustomElements().getValue(tag) + "\t";
				}
				builder = builder.addContents(details + "\n");
			}

			responseObserver.onNext(builder.build());
			responseObserver.onCompleted();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (ServiceException e) {
			e.printStackTrace();
		}
	}

	@Override
	public void fetchRowCol(FetchRowColRequest request, StreamObserver<FetchRowColReply> responseObserver) {
		String url = "https://spreadsheets.google.com/feeds/worksheets/" + request.getSpreadsheetId() + "/private/full";

		SpreadsheetService service = new SpreadsheetService("GoogleSheetsGrpc");
		FetchRowColReply.Builder builder = FetchRowColReply.newBuilder();

		String parms = null, details = null, del = null;
		try {
			GoogleCredential credential = getCredentials();
			service.setOAuth2Credentials(credential);

			URL WORKSHEET_FEED_URL = new URL(url);
			WorksheetFeed feed = service.getFeed(WORKSHEET_FEED_URL, WorksheetFeed.class);

			// Make a request to the API to fetch information about all
			// worksheets in the spreadsheet.
			List<WorksheetEntry> worksheets = feed.getEntries();

			WorksheetEntry worksheet = null;
			for (int i = 0; i < worksheets.size(); i++) {
				worksheet = worksheets.get(i);
				// Get the required worksheet.
				if (worksheet.getTitle().getPlainText().equals(request.getWorksheetName()))
					break;
			}

			if (request.getRow() >= 0) {
				parms = "?min-row=" + request.getRow() + "&max-row=" + request.getRow();
				del = "\t";
			} else if (request.getCol() >= 0) {
				parms = "?min-col=" + request.getCol() + "&max-col=" + request.getCol();
				del = "\n";
			}

			URL cellFeedUrl = new URL(worksheet.getCellFeedUrl().toString() + parms);
			CellFeed cellFeed = service.getFeed(cellFeedUrl, CellFeed.class);
			for (CellEntry cellEntry : cellFeed.getEntries()) {
				builder = builder.addContents(cellEntry.getCell().getValue() + del);
			}

			responseObserver.onNext(builder.build());
			responseObserver.onCompleted();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (ServiceException e) {
			e.printStackTrace();
		}
	}

	@Override
	public void deleteRow(DeleteRowRequest request, StreamObserver<DeleteRowReply> responseObserver) {
		String url = "https://spreadsheets.google.com/feeds/worksheets/" + request.getSpreadsheetId() + "/private/full";

		SpreadsheetService service = new SpreadsheetService("GoogleSheetsGrpc");
		DeleteRowReply reply = null;

		try {
			GoogleCredential credential = getCredentials();
			service.setOAuth2Credentials(credential);

			URL WORKSHEET_FEED_URL = new URL(url);
			WorksheetFeed feed = service.getFeed(WORKSHEET_FEED_URL, WorksheetFeed.class);

			// Make a request to the API to fetch information about all
			// worksheets in the spreadsheet.
			List<WorksheetEntry> worksheets = feed.getEntries();

			WorksheetEntry worksheet = null;
			for (int i = 0; i < worksheets.size(); i++) {
				worksheet = worksheets.get(i);
				// Get the required worksheet.
				if (worksheet.getTitle().getPlainText().equals(request.getWorksheetName()))
					break;
			}

			// Fetch the list feed of the worksheet.
			URL listFeedUrl = worksheet.getListFeedUrl();
			ListFeed listFeed = service.getFeed(listFeedUrl, ListFeed.class);

			ListEntry row = listFeed.getEntries().get(request.getRow());

			// Delete the row using the API.
			row.delete();

			reply.newBuilder().setSuccess("Row: " + row.getPlainTextContent() + " is deleted. Please check").build();

			responseObserver.onNext(reply);
			responseObserver.onCompleted();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (ServiceException e) {
			e.printStackTrace();
		}
	}

}
