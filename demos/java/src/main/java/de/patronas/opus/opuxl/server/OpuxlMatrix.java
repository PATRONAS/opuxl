/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.server;

import java.util.ArrayList;
import java.util.List;

/**
 * @author stepan
 */
public class OpuxlMatrix {

  private List<MatrixHeader> headers;
  private final List<List<? extends Object>> data = new ArrayList<>();

  /**
   * @return the data
   */
  public List<List<? extends Object>> getData() {
    return data;
  }

  /**
   * @return the header
   */
  public List<MatrixHeader> getHeaders() {
    return headers;
  }

  /**
   * @param header
   *          the header to set
   */
  public void setHeaders(final List<MatrixHeader> headers) {
    this.headers = headers;
  }

}
