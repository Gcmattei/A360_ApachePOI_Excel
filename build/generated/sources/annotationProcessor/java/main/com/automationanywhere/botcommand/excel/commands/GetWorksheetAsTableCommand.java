package com.automationanywhere.botcommand.excel.commands;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.BotCommand;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.automationanywhere.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import com.automationanywhere.toolchain.runtime.session.CloseableSessionObject;
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

public final class GetWorksheetAsTableCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(GetWorksheetAsTableCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    GetWorksheetAsTable command = new GetWorksheetAsTable();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("sheetOption") && parameters.get("sheetOption") != null && parameters.get("sheetOption").get() != null) {
      convertedParameters.put("sheetOption", parameters.get("sheetOption").get());
      if(convertedParameters.get("sheetOption") !=null && !(convertedParameters.get("sheetOption") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetOption", "String", parameters.get("sheetOption").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("sheetOption") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sheetOption"));
    }
    if(convertedParameters.get("sheetOption") != null) {
      switch((String)convertedParameters.get("sheetOption")) {
        case "Active" : {

        } break;
        case "Specific" : {
          if(parameters.containsKey("sheetName") && parameters.get("sheetName") != null && parameters.get("sheetName").get() != null) {
            convertedParameters.put("sheetName", parameters.get("sheetName").get());
            if(convertedParameters.get("sheetName") !=null && !(convertedParameters.get("sheetName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetName", "String", parameters.get("sheetName").get().getClass().getSimpleName()));
            }
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","sheetOption"));
      }
    }

    if(parameters.containsKey("areaOption") && parameters.get("areaOption") != null && parameters.get("areaOption").get() != null) {
      convertedParameters.put("areaOption", parameters.get("areaOption").get());
      if(convertedParameters.get("areaOption") !=null && !(convertedParameters.get("areaOption") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","areaOption", "String", parameters.get("areaOption").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("areaOption") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","areaOption"));
    }
    if(convertedParameters.get("areaOption") != null) {
      switch((String)convertedParameters.get("areaOption")) {
        case "Sheet" : {

        } break;
        case "Range" : {
          if(parameters.containsKey("rangeA1") && parameters.get("rangeA1") != null && parameters.get("rangeA1").get() != null) {
            convertedParameters.put("rangeA1", parameters.get("rangeA1").get());
            if(convertedParameters.get("rangeA1") !=null && !(convertedParameters.get("rangeA1") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","rangeA1", "String", parameters.get("rangeA1").get().getClass().getSimpleName()));
            }
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","areaOption"));
      }
    }

    if(parameters.containsKey("rowsOption") && parameters.get("rowsOption") != null && parameters.get("rowsOption").get() != null) {
      convertedParameters.put("rowsOption", parameters.get("rowsOption").get());
      if(convertedParameters.get("rowsOption") !=null && !(convertedParameters.get("rowsOption") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","rowsOption", "String", parameters.get("rowsOption").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("rowsOption") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","rowsOption"));
    }
    if(convertedParameters.get("rowsOption") != null) {
      switch((String)convertedParameters.get("rowsOption")) {
        case "All" : {

        } break;
        case "Range" : {
          if(parameters.containsKey("firstRowOneBased") && parameters.get("firstRowOneBased") != null && parameters.get("firstRowOneBased").get() != null) {
            convertedParameters.put("firstRowOneBased", parameters.get("firstRowOneBased").get());
            if(convertedParameters.get("firstRowOneBased") !=null && !(convertedParameters.get("firstRowOneBased") instanceof Double)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","firstRowOneBased", "Double", parameters.get("firstRowOneBased").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("firstRowOneBased") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","firstRowOneBased"));
          }

          if(parameters.containsKey("lastRowOneBased") && parameters.get("lastRowOneBased") != null && parameters.get("lastRowOneBased").get() != null) {
            convertedParameters.put("lastRowOneBased", parameters.get("lastRowOneBased").get());
            if(convertedParameters.get("lastRowOneBased") !=null && !(convertedParameters.get("lastRowOneBased") instanceof Double)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","lastRowOneBased", "Double", parameters.get("lastRowOneBased").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("lastRowOneBased") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","lastRowOneBased"));
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","rowsOption"));
      }
    }

    if(parameters.containsKey("hasHeader") && parameters.get("hasHeader") != null && parameters.get("hasHeader").get() != null) {
      convertedParameters.put("hasHeader", parameters.get("hasHeader").get());
      if(convertedParameters.get("hasHeader") !=null && !(convertedParameters.get("hasHeader") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","hasHeader", "Boolean", parameters.get("hasHeader").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("hasHeader") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","hasHeader"));
    }

    if(parameters.containsKey("readOption") && parameters.get("readOption") != null && parameters.get("readOption").get() != null) {
      convertedParameters.put("readOption", parameters.get("readOption").get());
      if(convertedParameters.get("readOption") !=null && !(convertedParameters.get("readOption") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","readOption", "String", parameters.get("readOption").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("readOption") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","readOption"));
    }
    if(convertedParameters.get("readOption") != null) {
      switch((String)convertedParameters.get("readOption")) {
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
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","readOption"));
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
      Optional<Value> result =  Optional.ofNullable(command.action((String)convertedParameters.get("sheetOption"),(String)convertedParameters.get("sheetName"),(String)convertedParameters.get("areaOption"),(String)convertedParameters.get("rangeA1"),(String)convertedParameters.get("rowsOption"),(Double)convertedParameters.get("firstRowOneBased"),(Double)convertedParameters.get("lastRowOneBased"),(Boolean)convertedParameters.get("hasHeader"),(String)convertedParameters.get("readOption"),(String)convertedParameters.get("visibleHelp"),(String)convertedParameters.get("valueHelp"),(WorkbookSession)convertedParameters.get("session")));
      return logger.traceExit(result);
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","action"));
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

  public Map<String, Value> executeAndReturnMany(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return null;
  }
}
