package com.davita.botcommand.excel.commands;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.BotCommand;
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

public final class FindTextInSheetCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(FindTextInSheetCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    FindTextInSheet command = new FindTextInSheet();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("fromKind") && parameters.get("fromKind") != null && parameters.get("fromKind").get() != null) {
      convertedParameters.put("fromKind", parameters.get("fromKind").get());
      if(convertedParameters.get("fromKind") !=null && !(convertedParameters.get("fromKind") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","fromKind", "String", parameters.get("fromKind").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("fromKind") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","fromKind"));
    }
    if(convertedParameters.get("fromKind") != null) {
      switch((String)convertedParameters.get("fromKind")) {
        case "BEGIN" : {

        } break;
        case "END" : {

        } break;
        case "ACTIVE" : {

        } break;
        case "SPECIFIC" : {
          if(parameters.containsKey("fromCellA1") && parameters.get("fromCellA1") != null && parameters.get("fromCellA1").get() != null) {
            convertedParameters.put("fromCellA1", parameters.get("fromCellA1").get());
            if(convertedParameters.get("fromCellA1") !=null && !(convertedParameters.get("fromCellA1") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","fromCellA1", "String", parameters.get("fromCellA1").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("fromCellA1") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","fromCellA1"));
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","fromKind"));
      }
    }

    if(parameters.containsKey("toKind") && parameters.get("toKind") != null && parameters.get("toKind").get() != null) {
      convertedParameters.put("toKind", parameters.get("toKind").get());
      if(convertedParameters.get("toKind") !=null && !(convertedParameters.get("toKind") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","toKind", "String", parameters.get("toKind").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("toKind") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","toKind"));
    }
    if(convertedParameters.get("toKind") != null) {
      switch((String)convertedParameters.get("toKind")) {
        case "BEGIN" : {

        } break;
        case "END" : {

        } break;
        case "ACTIVE" : {

        } break;
        case "SPECIFIC" : {
          if(parameters.containsKey("toCellA1") && parameters.get("toCellA1") != null && parameters.get("toCellA1").get() != null) {
            convertedParameters.put("toCellA1", parameters.get("toCellA1").get());
            if(convertedParameters.get("toCellA1") !=null && !(convertedParameters.get("toCellA1") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","toCellA1", "String", parameters.get("toCellA1").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("toCellA1") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","toCellA1"));
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","toKind"));
      }
    }

    if(parameters.containsKey("query") && parameters.get("query") != null && parameters.get("query").get() != null) {
      convertedParameters.put("query", parameters.get("query").get());
      if(convertedParameters.get("query") !=null && !(convertedParameters.get("query") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","query", "String", parameters.get("query").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("direction") && parameters.get("direction") != null && parameters.get("direction").get() != null) {
      convertedParameters.put("direction", parameters.get("direction").get());
      if(convertedParameters.get("direction") !=null && !(convertedParameters.get("direction") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","direction", "String", parameters.get("direction").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("direction") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","direction"));
    }
    if(convertedParameters.get("direction") != null) {
      switch((String)convertedParameters.get("direction")) {
        case "BY_ROW" : {

        } break;
        case "BY_COL" : {

        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","direction"));
      }
    }

    if(parameters.containsKey("matchCase") && parameters.get("matchCase") != null && parameters.get("matchCase").get() != null) {
      convertedParameters.put("matchCase", parameters.get("matchCase").get());
      if(convertedParameters.get("matchCase") !=null && !(convertedParameters.get("matchCase") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","matchCase", "Boolean", parameters.get("matchCase").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("matchCase") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","matchCase"));
    }

    if(parameters.containsKey("matchWhole") && parameters.get("matchWhole") != null && parameters.get("matchWhole").get() != null) {
      convertedParameters.put("matchWhole", parameters.get("matchWhole").get());
      if(convertedParameters.get("matchWhole") !=null && !(convertedParameters.get("matchWhole") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","matchWhole", "Boolean", parameters.get("matchWhole").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("matchWhole") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","matchWhole"));
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
      Optional<Value> result =  Optional.ofNullable(command.action((String)convertedParameters.get("fromKind"),(String)convertedParameters.get("fromCellA1"),(String)convertedParameters.get("toKind"),(String)convertedParameters.get("toCellA1"),(String)convertedParameters.get("query"),(String)convertedParameters.get("direction"),(Boolean)convertedParameters.get("matchCase"),(Boolean)convertedParameters.get("matchWhole"),(WorkbookSession)convertedParameters.get("session")));
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
