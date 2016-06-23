/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl;

import java.lang.reflect.InvocationTargetException;
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
class InstanceMethod {

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
  public InstanceMethod(final Object instance, final Method method) {
    this.instance = instance;
    this.method = method;
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
      final List<ExcelArgumentAttribute> parameterArguments) {
    this(instance, method);
    setParameterArguments(parameterArguments);
  }

  /**
   * TODO Rethink the approach of sending everything back as a Object-2d-Array.
   * @return The Argument Attributes as a collection of String arrays with a
   *         name and a description each.
   * @deprecated
   */
  @Deprecated
  public Collection<String[]> toJsonArrayString() {
    return Collections2
        .transform(getParameterArguments(), new Function<ExcelArgumentAttribute, String[]>() {
          @Override
          public String[] apply(final ExcelArgumentAttribute arg) {
            return new String[] {
                arg.getName(), arg.getDescription()
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
  public Object[][] execute(final List<Object> args) {
    final Object executed = internalExecute(args);
    return convert(executed);
  }

  private Object[][] convert(final Object executed) {
    final Object[][] result;

    // Opuxl always expects an object matrix as its result to be able to easily
    // insert it into the Excel sheet.
    if (executed instanceof Object[][]) {
      result = (Object[][]) executed;
    } else if (executed instanceof Object[]) {
      result = new Object[][] {
        (Object[]) executed
      };
    } else {
      result = new Object[][] {
        new Object[] {
          executed
        }
      };
    }

    return result;
  }

  private Object internalExecute(final List<Object> args) {
    try {
      return method.invoke(instance, args.toArray());
    } catch (IllegalAccessException | IllegalArgumentException
        | InvocationTargetException ex) {
      LOG.error("Error while executing method: " + method.getName() + " / "
          + ex.getMessage());
      throw new RuntimeException(ex);
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

}
