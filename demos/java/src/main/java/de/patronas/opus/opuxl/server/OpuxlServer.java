/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.server;

import java.io.IOException;
import java.net.ServerSocket;
import java.net.Socket;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * The Opuxl Server creates a server socket on the given port and handles all
 * incoming messages.
 * @author stepan
 */
public class OpuxlServer extends Thread {

  private static final Logger LOG = LoggerFactory.getLogger(OpuxlServer.class);

  /**
   * The Default Port.
   */
  public static final int DEFAULT_PORT = 61379;
  /**
   * The Port of the Socket Server.
   */
  private Integer port;

  private ServerSocket serverSocket;
  private final FunctionHandler functionHandler;

  /**
   * Constructor.
   * @param functionHandler
   *          the function handler with its registered methods
   */
  public OpuxlServer(final FunctionHandler functionHandler) {
    this.functionHandler = functionHandler;
  }

  /**
   * Constructor.
   * @param port
   *          the port for the socket server
   * @param functionHandler
   *          the function handler with its registered methods
   */
  public OpuxlServer(final int port, final FunctionHandler functionHandler) {
    this.functionHandler = functionHandler;
    this.port = port;
  }

  /**
   * Starts the server.
   */
  @Override
  public void run() {
    final int portToUse = port != null ? port : DEFAULT_PORT;
    try {
      serverSocket = new ServerSocket(portToUse);
      LOG.info("Opuxel Server Socket startet on port: "
          + serverSocket.getLocalPort());
    } catch (final IOException e) {
      LOG.error("Can't start server on port: " + portToUse);
    }

    while (serverSocket != null && !isInterrupted()) {
      try {
        final Socket socket = serverSocket.accept();
        new MessageDispatcher(socket, functionHandler).start();
      } catch (final IOException ex) {
        LOG.error("Error while processing client connection.");
        throw new RuntimeException(ex);
      }
    }

  }
}
