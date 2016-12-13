/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.server;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.base.Function;
import com.google.common.collect.Collections2;

/**
 * Wrapper class for a specific method of an object instance.
 * @author stepan
 */
public class InstanceMethod {

  private static final Logger LOG = LoggerFactory
      .getLogger(InstanceMethod.class);

  /*
   * The object on which the method is invoked.
   */
  private final Object instance;
  /*
   * The method which is invoked.
   */
  private final Method method;

  /*
   * The method description.
   */
  private final String description;
  /*
   * The description of each parameter.
   */
  private List<ExcelArgumentAttribute> parameterArguments = new ArrayList<>();

  /**
   * Constructor.
   * @param instance
   *          the object instance
   * @param method
   *          the method
   */
  public InstanceMethod(final Object instance,
      final Method method,
      final String description) {
    this.instance = instance;
    this.method = method;
    this.description = description;
  }

  /**
   * Constructor.
   * @param instance
   *          the object instance
   * @param method
   *          the method
   * @param parameterTypes
   *          the parameter types
   * @param parameterArguments
   *          the parameter arguments
   */
  public InstanceMethod(final Object instance,
      final Method method,
      final String description,
      final List<ExcelArgumentAttribute> parameterArguments) {
    this(instance, method, description);
    setParameterArguments(parameterArguments);
  }

  /**
   * @return The Argument Attributes as a collection of String arrays with a
   *         name and a description each.
   * @deprecated("Rethink the approach of sending everything back as a
   *                      Object-2d-Array. Rather use a serializable wrapper
   *                      class which is also present in the C# Code.")
   */
  @Deprecated
  public Collection<String[]> toJsonArrayString() {
    return Collections2
        .transform(getParameterArguments(), new Function<ExcelArgumentAttribute, String[]>() {
          @Override
          public String[] apply(final ExcelArgumentAttribute arg) {
            return new String[] {
                arg.getName(),
                arg.getDescription() + (arg.isOptional() ? " (Optional)" : ""),
                arg.getType().toString(),
                Boolean.toString(arg.isOptional())
            };
          }
        });
  }

  /**
   * Executes this method with the passed arguments.
   * @param args
   *          the arguments
   * @return the result of the function
   */
  public OpuxlMatrix execute(final List<Object> args) {
    final OpuxlMatrix executed = internalExecute(args);
    return executed;
  }

  private OpuxlMatrix internalExecute(final List<Object> args) {
    try {
      return (OpuxlMatrix) method.invoke(instance, args.toArray());
    } catch (final Exception ex) {
      LOG.error("Error while executing method: " + method.getName() + " / "
          + ex.getCause().getMessage());
      throw new RuntimeException(ex.getCause());
    }
  }

  /**
   * @return the parameterArguments
   */
  public List<ExcelArgumentAttribute> getParameterArguments() {
    return parameterArguments;
  }

  /**
   * @param parameterArguments
   *          the parameterArguments to set
   */
  public void setParameterArguments(final List<ExcelArgumentAttribute> parameterArguments) {
    this.parameterArguments = parameterArguments;
  }

  /**
   * @return the description
   */
  public String getDescription() {
    return description;
  }

}
