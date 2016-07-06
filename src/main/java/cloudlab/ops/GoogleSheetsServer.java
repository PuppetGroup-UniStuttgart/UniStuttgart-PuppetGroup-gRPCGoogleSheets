package cloudlab.ops;

import java.util.logging.Logger;

import cloudlab.GoogleSheetsProto.GoogleSheetsGrpc;
import io.grpc.Server;
import io.grpc.ServerBuilder;
/**
 * GoogleSheetsServer:
 * The service GoogleSheets is bound to this server's port
 */

public class GoogleSheetsServer {
	private static final Logger logger = Logger.getLogger(GoogleSheetsServer.class.getName());

	private int port = 50054;
	private Server server;

	private void start() throws Exception {
		server = ServerBuilder.forPort(port).addService(GoogleSheetsGrpc.bindService(new GoogleSheetsImpl())).build().start();
		logger.info("Server started, listening on " + port);
		Runtime.getRuntime().addShutdownHook(new Thread() {
			@Override
			public void run() {
				System.err.println("*** shutting down gRPC server since JVM is shutting down");
				GoogleSheetsServer.this.stop();
				System.err.println("*** server shut down");
			}
		});
	}

	private void stop() {
		if (server != null) {
			server.shutdown();
		}
	}

/*	 Await termination on the main thread since the grpc library uses daemon
	 threads.*/

	private void blockUntilShutdown() throws InterruptedException {
		if (server != null) {
			server.awaitTermination();
		}
	}

	public static void main(String[] args) throws Exception {
		final GoogleSheetsServer server = new GoogleSheetsServer();
		server.start();
		server.blockUntilShutdown();
	}
}
