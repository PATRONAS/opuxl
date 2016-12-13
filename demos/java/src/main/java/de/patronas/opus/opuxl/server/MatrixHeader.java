/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.server;

/**
 * @author stepan
 */
public class MatrixHeader {

  private String text;
  private ExcelType type;

  /**
   * Constructor.
   * @param text
   *          the title of the header
   * @param type
   *          the type of the entries
   */
  public MatrixHeader(final String text, final ExcelType type) {
    this.text = text;
    this.type = type;
  }

  /**
   * @return the text
   */
  public String getText() {
    return text;
  }

  /**
   * @param text
   *          the text to set
   */
  public void setText(final String text) {
    this.text = text;
  }

  /**
   * @return the type
   */
  public ExcelType getType() {
    return type;
  }

  /**
   * @param type
   *          the type to set
   */
  public void setType(final ExcelType type) {
    this.type = type;
  }
}
