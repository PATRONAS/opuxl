/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.server;

/**
 * @author stepan
 */
public enum ExcelType {

  CURRENCY("currency"), // SQL Numerical
  DATETIME("datetime"), // SQL Timestamp
  LOGICAL("logical"), // SQL Bit
  NUMBER("number"), // SQL Double
  TEXT("text"); // SQL Varchar

  private final String type;

  private ExcelType(final String type) {
    this.type = type;
  }

  public String get() {
    return type;
  }
}
