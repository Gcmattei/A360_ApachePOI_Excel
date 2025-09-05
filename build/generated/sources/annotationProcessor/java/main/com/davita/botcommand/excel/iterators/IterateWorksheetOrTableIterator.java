package com.davita.botcommand.excel.iterators;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.CommandIterator;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import com.automationanywhere.toolchain.runtime.session.CloseableSessionObject;
import com.davita.botcommand.excel.sessions.WorkbookSession;
import java.lang.Boolean;
import java.lang.ClassCastException;
import java.lang.Deprecated;
import java.lang.Double;
import java.lang.Object;
import java.lang.String;
import java.lang.Throwable;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public final class IterateWorksheetOrTableIterator implements CommandIterator {
  private static final Logger logger = LogManager.getLogger(IterateWorksheetOrTableIterator.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  private final IterateWorksheetOrTable command = new IterateWorksheetOrTable();

  public boolean hasNext(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return hasNext(null, parameters, sessionMap, null);
  }

  public boolean hasNext(GlobalSessionContext globalSessionContext, Map<String, Value> parameters,
      Map<String, Object> sessionMap, Map<String, Value> packageSettings) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null);
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("mode") && parameters.get("mode") != null && parameters.get("mode").get() != null) {
      convertedParameters.put("mode", parameters.get("mode").get());
      if(convertedParameters.get("mode") !=null && !(convertedParameters.get("mode") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","mode", "String", parameters.get("mode").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("mode") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","mode"));
    }
    try {
      command.setMode( (String)convertedParameters.get("mode"));
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","mode", "String", parameters.get("mode") != null ? (parameters.get("mode").get() != null ? parameters.get("mode").get().getClass().toString() : "null") : "null"),e);
    }
    if(convertedParameters.get("mode") != null) {
      switch((String)convertedParameters.get("mode")) {
        case "TABLE" : {
          if(parameters.containsKey("tableName") && parameters.get("tableName") != null && parameters.get("tableName").get() != null) {
            convertedParameters.put("tableName", parameters.get("tableName").get());
            if(convertedParameters.get("tableName") !=null && !(convertedParameters.get("tableName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableName", "String", parameters.get("tableName").get().getClass().getSimpleName()));
            }
          }
          try {
            command.setTableName( (String)convertedParameters.get("tableName"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","tableName", "String", parameters.get("tableName") != null ? (parameters.get("tableName").get() != null ? parameters.get("tableName").get().getClass().toString() : "null") : "null"),e);
          }

          if(parameters.containsKey("tableRowsMode") && parameters.get("tableRowsMode") != null && parameters.get("tableRowsMode").get() != null) {
            convertedParameters.put("tableRowsMode", parameters.get("tableRowsMode").get());
            if(convertedParameters.get("tableRowsMode") !=null && !(convertedParameters.get("tableRowsMode") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableRowsMode", "String", parameters.get("tableRowsMode").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("tableRowsMode") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","tableRowsMode"));
          }
          try {
            command.setTableRowsMode( (String)convertedParameters.get("tableRowsMode"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","tableRowsMode", "String", parameters.get("tableRowsMode") != null ? (parameters.get("tableRowsMode").get() != null ? parameters.get("tableRowsMode").get().getClass().toString() : "null") : "null"),e);
          }
          if(convertedParameters.get("tableRowsMode") != null) {
            switch((String)convertedParameters.get("tableRowsMode")) {
              case "ALL_ROWS" : {

              } break;
              case "SPECIFIC_ROWS" : {
                if(parameters.containsKey("tableStartRowOneBased") && parameters.get("tableStartRowOneBased") != null && parameters.get("tableStartRowOneBased").get() != null) {
                  convertedParameters.put("tableStartRowOneBased", parameters.get("tableStartRowOneBased").get());
                  if(convertedParameters.get("tableStartRowOneBased") !=null && !(convertedParameters.get("tableStartRowOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableStartRowOneBased", "Double", parameters.get("tableStartRowOneBased").get().getClass().getSimpleName()));
                  }
                }
                try {
                  command.setTableStartRowOneBased( (Double)convertedParameters.get("tableStartRowOneBased"));
                }
                catch (ClassCastException e) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","tableStartRowOneBased", "Double", parameters.get("tableStartRowOneBased") != null ? (parameters.get("tableStartRowOneBased").get() != null ? parameters.get("tableStartRowOneBased").get().getClass().toString() : "null") : "null"),e);
                }

                if(parameters.containsKey("tableEndRowOneBased") && parameters.get("tableEndRowOneBased") != null && parameters.get("tableEndRowOneBased").get() != null) {
                  convertedParameters.put("tableEndRowOneBased", parameters.get("tableEndRowOneBased").get());
                  if(convertedParameters.get("tableEndRowOneBased") !=null && !(convertedParameters.get("tableEndRowOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableEndRowOneBased", "Double", parameters.get("tableEndRowOneBased").get().getClass().getSimpleName()));
                  }
                }
                try {
                  command.setTableEndRowOneBased( (Double)convertedParameters.get("tableEndRowOneBased"));
                }
                catch (ClassCastException e) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","tableEndRowOneBased", "Double", parameters.get("tableEndRowOneBased") != null ? (parameters.get("tableEndRowOneBased").get() != null ? parameters.get("tableEndRowOneBased").get().getClass().toString() : "null") : "null"),e);
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","tableRowsMode"));
            }
          }


        } break;
        case "WORKSHEET" : {
          if(parameters.containsKey("sheetName") && parameters.get("sheetName") != null && parameters.get("sheetName").get() != null) {
            convertedParameters.put("sheetName", parameters.get("sheetName").get());
            if(convertedParameters.get("sheetName") !=null && !(convertedParameters.get("sheetName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetName", "String", parameters.get("sheetName").get().getClass().getSimpleName()));
            }
          }
          try {
            command.setSheetName( (String)convertedParameters.get("sheetName"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","sheetName", "String", parameters.get("sheetName") != null ? (parameters.get("sheetName").get() != null ? parameters.get("sheetName").get().getClass().toString() : "null") : "null"),e);
          }

          if(parameters.containsKey("wsRangeMode") && parameters.get("wsRangeMode") != null && parameters.get("wsRangeMode").get() != null) {
            convertedParameters.put("wsRangeMode", parameters.get("wsRangeMode").get());
            if(convertedParameters.get("wsRangeMode") !=null && !(convertedParameters.get("wsRangeMode") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsRangeMode", "String", parameters.get("wsRangeMode").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("wsRangeMode") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsRangeMode"));
          }
          try {
            command.setWsRangeMode( (String)convertedParameters.get("wsRangeMode"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsRangeMode", "String", parameters.get("wsRangeMode") != null ? (parameters.get("wsRangeMode").get() != null ? parameters.get("wsRangeMode").get().getClass().toString() : "null") : "null"),e);
          }
          if(convertedParameters.get("wsRangeMode") != null) {
            switch((String)convertedParameters.get("wsRangeMode")) {
              case "ALL" : {

              } break;
              case "SPECIFIC" : {
                if(parameters.containsKey("wsRangeA1") && parameters.get("wsRangeA1") != null && parameters.get("wsRangeA1").get() != null) {
                  convertedParameters.put("wsRangeA1", parameters.get("wsRangeA1").get());
                  if(convertedParameters.get("wsRangeA1") !=null && !(convertedParameters.get("wsRangeA1") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsRangeA1", "String", parameters.get("wsRangeA1").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("wsRangeA1") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsRangeA1"));
                }
                try {
                  command.setWsRangeA1( (String)convertedParameters.get("wsRangeA1"));
                }
                catch (ClassCastException e) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsRangeA1", "String", parameters.get("wsRangeA1") != null ? (parameters.get("wsRangeA1").get() != null ? parameters.get("wsRangeA1").get().getClass().toString() : "null") : "null"),e);
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","wsRangeMode"));
            }
          }

          if(parameters.containsKey("wsRowsMode") && parameters.get("wsRowsMode") != null && parameters.get("wsRowsMode").get() != null) {
            convertedParameters.put("wsRowsMode", parameters.get("wsRowsMode").get());
            if(convertedParameters.get("wsRowsMode") !=null && !(convertedParameters.get("wsRowsMode") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsRowsMode", "String", parameters.get("wsRowsMode").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("wsRowsMode") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsRowsMode"));
          }
          try {
            command.setWsRowsMode( (String)convertedParameters.get("wsRowsMode"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsRowsMode", "String", parameters.get("wsRowsMode") != null ? (parameters.get("wsRowsMode").get() != null ? parameters.get("wsRowsMode").get().getClass().toString() : "null") : "null"),e);
          }
          if(convertedParameters.get("wsRowsMode") != null) {
            switch((String)convertedParameters.get("wsRowsMode")) {
              case "ALL_ROWS" : {

              } break;
              case "SPECIFIC_ROWS" : {
                if(parameters.containsKey("wsStartRowOneBased") && parameters.get("wsStartRowOneBased") != null && parameters.get("wsStartRowOneBased").get() != null) {
                  convertedParameters.put("wsStartRowOneBased", parameters.get("wsStartRowOneBased").get());
                  if(convertedParameters.get("wsStartRowOneBased") !=null && !(convertedParameters.get("wsStartRowOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsStartRowOneBased", "Double", parameters.get("wsStartRowOneBased").get().getClass().getSimpleName()));
                  }
                }
                try {
                  command.setWsStartRowOneBased( (Double)convertedParameters.get("wsStartRowOneBased"));
                }
                catch (ClassCastException e) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsStartRowOneBased", "Double", parameters.get("wsStartRowOneBased") != null ? (parameters.get("wsStartRowOneBased").get() != null ? parameters.get("wsStartRowOneBased").get().getClass().toString() : "null") : "null"),e);
                }

                if(parameters.containsKey("wsEndRowOneBased") && parameters.get("wsEndRowOneBased") != null && parameters.get("wsEndRowOneBased").get() != null) {
                  convertedParameters.put("wsEndRowOneBased", parameters.get("wsEndRowOneBased").get());
                  if(convertedParameters.get("wsEndRowOneBased") !=null && !(convertedParameters.get("wsEndRowOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsEndRowOneBased", "Double", parameters.get("wsEndRowOneBased").get().getClass().getSimpleName()));
                  }
                }
                try {
                  command.setWsEndRowOneBased( (Double)convertedParameters.get("wsEndRowOneBased"));
                }
                catch (ClassCastException e) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsEndRowOneBased", "Double", parameters.get("wsEndRowOneBased") != null ? (parameters.get("wsEndRowOneBased").get() != null ? parameters.get("wsEndRowOneBased").get().getClass().toString() : "null") : "null"),e);
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","wsRowsMode"));
            }
          }

          if(parameters.containsKey("wsHasHeader") && parameters.get("wsHasHeader") != null && parameters.get("wsHasHeader").get() != null) {
            convertedParameters.put("wsHasHeader", parameters.get("wsHasHeader").get());
            if(convertedParameters.get("wsHasHeader") !=null && !(convertedParameters.get("wsHasHeader") instanceof Boolean)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsHasHeader", "Boolean", parameters.get("wsHasHeader").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("wsHasHeader") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsHasHeader"));
          }
          try {
            command.setWsHasHeader( (Boolean)convertedParameters.get("wsHasHeader"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsHasHeader", "Boolean", parameters.get("wsHasHeader") != null ? (parameters.get("wsHasHeader").get() != null ? parameters.get("wsHasHeader").get().getClass().toString() : "null") : "null"),e);
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","mode"));
      }
    }

    if(parameters.containsKey("readMode") && parameters.get("readMode") != null && parameters.get("readMode").get() != null) {
      convertedParameters.put("readMode", parameters.get("readMode").get());
      if(convertedParameters.get("readMode") !=null && !(convertedParameters.get("readMode") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","readMode", "String", parameters.get("readMode").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("readMode") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","readMode"));
    }
    try {
      command.setReadMode( (String)convertedParameters.get("readMode"));
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","readMode", "String", parameters.get("readMode") != null ? (parameters.get("readMode").get() != null ? parameters.get("readMode").get().getClass().toString() : "null") : "null"),e);
    }
    if(convertedParameters.get("readMode") != null) {
      switch((String)convertedParameters.get("readMode")) {
        case "visible" : {
          if(parameters.containsKey("visibleHelp") && parameters.get("visibleHelp") != null && parameters.get("visibleHelp").get() != null) {
            convertedParameters.put("visibleHelp", parameters.get("visibleHelp").get());
            if(convertedParameters.get("visibleHelp") !=null && !(convertedParameters.get("visibleHelp") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","visibleHelp", "String", parameters.get("visibleHelp").get().getClass().getSimpleName()));
            }
          }


        } break;
        case "value" : {
          if(parameters.containsKey("valueHelp") && parameters.get("valueHelp") != null && parameters.get("valueHelp").get() != null) {
            convertedParameters.put("valueHelp", parameters.get("valueHelp").get());
            if(convertedParameters.get("valueHelp") !=null && !(convertedParameters.get("valueHelp") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","valueHelp", "String", parameters.get("valueHelp").get().getClass().getSimpleName()));
            }
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","readMode"));
      }
    }

    if(parameters.containsKey("session") && parameters.get("session") != null) {
      if(((SessionValue)parameters.get("session")).hasObjectValue()) {
        convertedParameters.put("session", ((SessionValue)parameters.get("session")).getSession());
      }
      else if(parameters.get("session").get() != null) {
        convertedParameters.put("session", sessionMap.get(parameters.get("session").get()));
      }
      if(convertedParameters.get("session")!=null && ((CloseableSessionObject)convertedParameters.get("session")).isClosed()) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("session.IsAlreadyClosed",parameters.get("session").get()));
      }
      if(convertedParameters.get("session") !=null && !(convertedParameters.get("session") instanceof WorkbookSession)) {
        Class[] interfaces = convertedParameters.get("session").getClass().getInterfaces();
        for (Class iface : interfaces) {
          if(iface.getName().equals(WorkbookSession.class.getCanonicalName())) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("session.package.mismatch","session" ));
          }
        }
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","session", "WorkbookSession", convertedParameters.get("session").getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("session") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","session"));
    }
    try {
      command.setSession( (WorkbookSession)convertedParameters.get("session"));
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","session", "WorkbookSession", parameters.get("session") != null ? (parameters.get("session").get() != null ? parameters.get("session").get().getClass().toString() : "null") : "null"),e);
    }

    try {
      boolean result = command.hasNext();
      logger.traceExit(result);
      return result;
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","hasNext"));
    }
    catch (BotCommandException e) {
      logger.fatal(e.getMessage(),e);
      throw e;
    }
    catch (Throwable e) {
      logger.fatal(e.getMessage(),e);
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.NotBotCommandException",e.getMessage()),e);
    }
  }

  @Deprecated
  public Optional<Map<String, Value>> next(Map<String, Value> parameters,
      Map<String, Object> sessionMap) {
    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.DepricatedNextMethod"));
  }

  public Optional<Value> nextOne(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return nextOne(null, parameters, sessionMap, null);
  }

  public Optional<Value> nextOne(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap,
      Map<String, Value> packageSettings) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null);
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("mode") && parameters.get("mode") != null && parameters.get("mode").get() != null) {
      convertedParameters.put("mode", parameters.get("mode").get());
      if(convertedParameters.get("mode") !=null && !(convertedParameters.get("mode") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","mode", "String", parameters.get("mode").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("mode") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","mode"));
    }
    try {
      command.setMode( (String)convertedParameters.get("mode"));
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","mode", "String", parameters.get("mode") != null ? (parameters.get("mode").get() != null ? parameters.get("mode").get().getClass().toString() : "null") : "null"),e);
    }
    if(convertedParameters.get("mode") != null) {
      switch((String)convertedParameters.get("mode")) {
        case "TABLE" : {
          if(parameters.containsKey("tableName") && parameters.get("tableName") != null && parameters.get("tableName").get() != null) {
            convertedParameters.put("tableName", parameters.get("tableName").get());
            if(convertedParameters.get("tableName") !=null && !(convertedParameters.get("tableName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableName", "String", parameters.get("tableName").get().getClass().getSimpleName()));
            }
          }
          try {
            command.setTableName( (String)convertedParameters.get("tableName"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","tableName", "String", parameters.get("tableName") != null ? (parameters.get("tableName").get() != null ? parameters.get("tableName").get().getClass().toString() : "null") : "null"),e);
          }

          if(parameters.containsKey("tableRowsMode") && parameters.get("tableRowsMode") != null && parameters.get("tableRowsMode").get() != null) {
            convertedParameters.put("tableRowsMode", parameters.get("tableRowsMode").get());
            if(convertedParameters.get("tableRowsMode") !=null && !(convertedParameters.get("tableRowsMode") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableRowsMode", "String", parameters.get("tableRowsMode").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("tableRowsMode") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","tableRowsMode"));
          }
          try {
            command.setTableRowsMode( (String)convertedParameters.get("tableRowsMode"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","tableRowsMode", "String", parameters.get("tableRowsMode") != null ? (parameters.get("tableRowsMode").get() != null ? parameters.get("tableRowsMode").get().getClass().toString() : "null") : "null"),e);
          }
          if(convertedParameters.get("tableRowsMode") != null) {
            switch((String)convertedParameters.get("tableRowsMode")) {
              case "ALL_ROWS" : {

              } break;
              case "SPECIFIC_ROWS" : {
                if(parameters.containsKey("tableStartRowOneBased") && parameters.get("tableStartRowOneBased") != null && parameters.get("tableStartRowOneBased").get() != null) {
                  convertedParameters.put("tableStartRowOneBased", parameters.get("tableStartRowOneBased").get());
                  if(convertedParameters.get("tableStartRowOneBased") !=null && !(convertedParameters.get("tableStartRowOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableStartRowOneBased", "Double", parameters.get("tableStartRowOneBased").get().getClass().getSimpleName()));
                  }
                }
                try {
                  command.setTableStartRowOneBased( (Double)convertedParameters.get("tableStartRowOneBased"));
                }
                catch (ClassCastException e) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","tableStartRowOneBased", "Double", parameters.get("tableStartRowOneBased") != null ? (parameters.get("tableStartRowOneBased").get() != null ? parameters.get("tableStartRowOneBased").get().getClass().toString() : "null") : "null"),e);
                }

                if(parameters.containsKey("tableEndRowOneBased") && parameters.get("tableEndRowOneBased") != null && parameters.get("tableEndRowOneBased").get() != null) {
                  convertedParameters.put("tableEndRowOneBased", parameters.get("tableEndRowOneBased").get());
                  if(convertedParameters.get("tableEndRowOneBased") !=null && !(convertedParameters.get("tableEndRowOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableEndRowOneBased", "Double", parameters.get("tableEndRowOneBased").get().getClass().getSimpleName()));
                  }
                }
                try {
                  command.setTableEndRowOneBased( (Double)convertedParameters.get("tableEndRowOneBased"));
                }
                catch (ClassCastException e) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","tableEndRowOneBased", "Double", parameters.get("tableEndRowOneBased") != null ? (parameters.get("tableEndRowOneBased").get() != null ? parameters.get("tableEndRowOneBased").get().getClass().toString() : "null") : "null"),e);
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","tableRowsMode"));
            }
          }


        } break;
        case "WORKSHEET" : {
          if(parameters.containsKey("sheetName") && parameters.get("sheetName") != null && parameters.get("sheetName").get() != null) {
            convertedParameters.put("sheetName", parameters.get("sheetName").get());
            if(convertedParameters.get("sheetName") !=null && !(convertedParameters.get("sheetName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetName", "String", parameters.get("sheetName").get().getClass().getSimpleName()));
            }
          }
          try {
            command.setSheetName( (String)convertedParameters.get("sheetName"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","sheetName", "String", parameters.get("sheetName") != null ? (parameters.get("sheetName").get() != null ? parameters.get("sheetName").get().getClass().toString() : "null") : "null"),e);
          }

          if(parameters.containsKey("wsRangeMode") && parameters.get("wsRangeMode") != null && parameters.get("wsRangeMode").get() != null) {
            convertedParameters.put("wsRangeMode", parameters.get("wsRangeMode").get());
            if(convertedParameters.get("wsRangeMode") !=null && !(convertedParameters.get("wsRangeMode") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsRangeMode", "String", parameters.get("wsRangeMode").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("wsRangeMode") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsRangeMode"));
          }
          try {
            command.setWsRangeMode( (String)convertedParameters.get("wsRangeMode"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsRangeMode", "String", parameters.get("wsRangeMode") != null ? (parameters.get("wsRangeMode").get() != null ? parameters.get("wsRangeMode").get().getClass().toString() : "null") : "null"),e);
          }
          if(convertedParameters.get("wsRangeMode") != null) {
            switch((String)convertedParameters.get("wsRangeMode")) {
              case "ALL" : {

              } break;
              case "SPECIFIC" : {
                if(parameters.containsKey("wsRangeA1") && parameters.get("wsRangeA1") != null && parameters.get("wsRangeA1").get() != null) {
                  convertedParameters.put("wsRangeA1", parameters.get("wsRangeA1").get());
                  if(convertedParameters.get("wsRangeA1") !=null && !(convertedParameters.get("wsRangeA1") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsRangeA1", "String", parameters.get("wsRangeA1").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("wsRangeA1") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsRangeA1"));
                }
                try {
                  command.setWsRangeA1( (String)convertedParameters.get("wsRangeA1"));
                }
                catch (ClassCastException e) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsRangeA1", "String", parameters.get("wsRangeA1") != null ? (parameters.get("wsRangeA1").get() != null ? parameters.get("wsRangeA1").get().getClass().toString() : "null") : "null"),e);
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","wsRangeMode"));
            }
          }

          if(parameters.containsKey("wsRowsMode") && parameters.get("wsRowsMode") != null && parameters.get("wsRowsMode").get() != null) {
            convertedParameters.put("wsRowsMode", parameters.get("wsRowsMode").get());
            if(convertedParameters.get("wsRowsMode") !=null && !(convertedParameters.get("wsRowsMode") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsRowsMode", "String", parameters.get("wsRowsMode").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("wsRowsMode") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsRowsMode"));
          }
          try {
            command.setWsRowsMode( (String)convertedParameters.get("wsRowsMode"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsRowsMode", "String", parameters.get("wsRowsMode") != null ? (parameters.get("wsRowsMode").get() != null ? parameters.get("wsRowsMode").get().getClass().toString() : "null") : "null"),e);
          }
          if(convertedParameters.get("wsRowsMode") != null) {
            switch((String)convertedParameters.get("wsRowsMode")) {
              case "ALL_ROWS" : {

              } break;
              case "SPECIFIC_ROWS" : {
                if(parameters.containsKey("wsStartRowOneBased") && parameters.get("wsStartRowOneBased") != null && parameters.get("wsStartRowOneBased").get() != null) {
                  convertedParameters.put("wsStartRowOneBased", parameters.get("wsStartRowOneBased").get());
                  if(convertedParameters.get("wsStartRowOneBased") !=null && !(convertedParameters.get("wsStartRowOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsStartRowOneBased", "Double", parameters.get("wsStartRowOneBased").get().getClass().getSimpleName()));
                  }
                }
                try {
                  command.setWsStartRowOneBased( (Double)convertedParameters.get("wsStartRowOneBased"));
                }
                catch (ClassCastException e) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsStartRowOneBased", "Double", parameters.get("wsStartRowOneBased") != null ? (parameters.get("wsStartRowOneBased").get() != null ? parameters.get("wsStartRowOneBased").get().getClass().toString() : "null") : "null"),e);
                }

                if(parameters.containsKey("wsEndRowOneBased") && parameters.get("wsEndRowOneBased") != null && parameters.get("wsEndRowOneBased").get() != null) {
                  convertedParameters.put("wsEndRowOneBased", parameters.get("wsEndRowOneBased").get());
                  if(convertedParameters.get("wsEndRowOneBased") !=null && !(convertedParameters.get("wsEndRowOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsEndRowOneBased", "Double", parameters.get("wsEndRowOneBased").get().getClass().getSimpleName()));
                  }
                }
                try {
                  command.setWsEndRowOneBased( (Double)convertedParameters.get("wsEndRowOneBased"));
                }
                catch (ClassCastException e) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsEndRowOneBased", "Double", parameters.get("wsEndRowOneBased") != null ? (parameters.get("wsEndRowOneBased").get() != null ? parameters.get("wsEndRowOneBased").get().getClass().toString() : "null") : "null"),e);
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","wsRowsMode"));
            }
          }

          if(parameters.containsKey("wsHasHeader") && parameters.get("wsHasHeader") != null && parameters.get("wsHasHeader").get() != null) {
            convertedParameters.put("wsHasHeader", parameters.get("wsHasHeader").get());
            if(convertedParameters.get("wsHasHeader") !=null && !(convertedParameters.get("wsHasHeader") instanceof Boolean)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsHasHeader", "Boolean", parameters.get("wsHasHeader").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("wsHasHeader") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsHasHeader"));
          }
          try {
            command.setWsHasHeader( (Boolean)convertedParameters.get("wsHasHeader"));
          }
          catch (ClassCastException e) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","wsHasHeader", "Boolean", parameters.get("wsHasHeader") != null ? (parameters.get("wsHasHeader").get() != null ? parameters.get("wsHasHeader").get().getClass().toString() : "null") : "null"),e);
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","mode"));
      }
    }

    if(parameters.containsKey("readMode") && parameters.get("readMode") != null && parameters.get("readMode").get() != null) {
      convertedParameters.put("readMode", parameters.get("readMode").get());
      if(convertedParameters.get("readMode") !=null && !(convertedParameters.get("readMode") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","readMode", "String", parameters.get("readMode").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("readMode") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","readMode"));
    }
    try {
      command.setReadMode( (String)convertedParameters.get("readMode"));
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","readMode", "String", parameters.get("readMode") != null ? (parameters.get("readMode").get() != null ? parameters.get("readMode").get().getClass().toString() : "null") : "null"),e);
    }
    if(convertedParameters.get("readMode") != null) {
      switch((String)convertedParameters.get("readMode")) {
        case "visible" : {
          if(parameters.containsKey("visibleHelp") && parameters.get("visibleHelp") != null && parameters.get("visibleHelp").get() != null) {
            convertedParameters.put("visibleHelp", parameters.get("visibleHelp").get());
            if(convertedParameters.get("visibleHelp") !=null && !(convertedParameters.get("visibleHelp") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","visibleHelp", "String", parameters.get("visibleHelp").get().getClass().getSimpleName()));
            }
          }


        } break;
        case "value" : {
          if(parameters.containsKey("valueHelp") && parameters.get("valueHelp") != null && parameters.get("valueHelp").get() != null) {
            convertedParameters.put("valueHelp", parameters.get("valueHelp").get());
            if(convertedParameters.get("valueHelp") !=null && !(convertedParameters.get("valueHelp") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","valueHelp", "String", parameters.get("valueHelp").get().getClass().getSimpleName()));
            }
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","readMode"));
      }
    }

    if(parameters.containsKey("session") && parameters.get("session") != null) {
      if(((SessionValue)parameters.get("session")).hasObjectValue()) {
        convertedParameters.put("session", ((SessionValue)parameters.get("session")).getSession());
      }
      else if(parameters.get("session").get() != null) {
        convertedParameters.put("session", sessionMap.get(parameters.get("session").get()));
      }
      if(convertedParameters.get("session")!=null && ((CloseableSessionObject)convertedParameters.get("session")).isClosed()) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("session.IsAlreadyClosed",parameters.get("session").get()));
      }
      if(convertedParameters.get("session") !=null && !(convertedParameters.get("session") instanceof WorkbookSession)) {
        Class[] interfaces = convertedParameters.get("session").getClass().getInterfaces();
        for (Class iface : interfaces) {
          if(iface.getName().equals(WorkbookSession.class.getCanonicalName())) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("session.package.mismatch","session" ));
          }
        }
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","session", "WorkbookSession", convertedParameters.get("session").getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("session") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","session"));
    }
    try {
      command.setSession( (WorkbookSession)convertedParameters.get("session"));
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.ClassCastException","session", "WorkbookSession", parameters.get("session") != null ? (parameters.get("session").get() != null ? parameters.get("session").get().getClass().toString() : "null") : "null"),e);
    }

    try {
      Optional<Value> result =  Optional.ofNullable(command.next());
      logger.traceExit(result);
      return result;
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","next"));
    }
    catch (BotCommandException e) {
      logger.fatal(e.getMessage(),e);
      throw e;
    }
    catch (Throwable e) {
      logger.fatal(e.getMessage(),e);
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.NotBotCommandException",e.getMessage()),e);
    }
  }
}
