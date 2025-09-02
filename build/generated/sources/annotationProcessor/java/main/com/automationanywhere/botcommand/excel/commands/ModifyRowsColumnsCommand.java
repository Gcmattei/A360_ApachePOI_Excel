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

public final class ModifyRowsColumnsCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(ModifyRowsColumnsCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    ModifyRowsColumns command = new ModifyRowsColumns();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("group") && parameters.get("group") != null && parameters.get("group").get() != null) {
      convertedParameters.put("group", parameters.get("group").get());
      if(convertedParameters.get("group") !=null && !(convertedParameters.get("group") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","group", "String", parameters.get("group").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("group") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","group"));
    }
    if(convertedParameters.get("group") != null) {
      switch((String)convertedParameters.get("group")) {
        case "ROWS" : {
          if(parameters.containsKey("rowOperation") && parameters.get("rowOperation") != null && parameters.get("rowOperation").get() != null) {
            convertedParameters.put("rowOperation", parameters.get("rowOperation").get());
            if(convertedParameters.get("rowOperation") !=null && !(convertedParameters.get("rowOperation") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","rowOperation", "String", parameters.get("rowOperation").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("rowOperation") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","rowOperation"));
          }
          if(convertedParameters.get("rowOperation") != null) {
            switch((String)convertedParameters.get("rowOperation")) {
              case "INSERT" : {
                if(parameters.containsKey("rowsTargetInsert") && parameters.get("rowsTargetInsert") != null && parameters.get("rowsTargetInsert").get() != null) {
                  convertedParameters.put("rowsTargetInsert", parameters.get("rowsTargetInsert").get());
                  if(convertedParameters.get("rowsTargetInsert") !=null && !(convertedParameters.get("rowsTargetInsert") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","rowsTargetInsert", "String", parameters.get("rowsTargetInsert").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("rowsTargetInsert") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","rowsTargetInsert"));
                }


              } break;
              case "DELETE" : {
                if(parameters.containsKey("rowsTargetDelete") && parameters.get("rowsTargetDelete") != null && parameters.get("rowsTargetDelete").get() != null) {
                  convertedParameters.put("rowsTargetDelete", parameters.get("rowsTargetDelete").get());
                  if(convertedParameters.get("rowsTargetDelete") !=null && !(convertedParameters.get("rowsTargetDelete") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","rowsTargetDelete", "String", parameters.get("rowsTargetDelete").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("rowsTargetDelete") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","rowsTargetDelete"));
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","rowOperation"));
            }
          }


        } break;
        case "COLUMNS" : {
          if(parameters.containsKey("colOperation") && parameters.get("colOperation") != null && parameters.get("colOperation").get() != null) {
            convertedParameters.put("colOperation", parameters.get("colOperation").get());
            if(convertedParameters.get("colOperation") !=null && !(convertedParameters.get("colOperation") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","colOperation", "String", parameters.get("colOperation").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("colOperation") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","colOperation"));
          }
          if(convertedParameters.get("colOperation") != null) {
            switch((String)convertedParameters.get("colOperation")) {
              case "INSERT" : {
                if(parameters.containsKey("colsTargetInsert") && parameters.get("colsTargetInsert") != null && parameters.get("colsTargetInsert").get() != null) {
                  convertedParameters.put("colsTargetInsert", parameters.get("colsTargetInsert").get());
                  if(convertedParameters.get("colsTargetInsert") !=null && !(convertedParameters.get("colsTargetInsert") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","colsTargetInsert", "String", parameters.get("colsTargetInsert").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("colsTargetInsert") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","colsTargetInsert"));
                }


              } break;
              case "DELETE" : {
                if(parameters.containsKey("colsTargetDelete") && parameters.get("colsTargetDelete") != null && parameters.get("colsTargetDelete").get() != null) {
                  convertedParameters.put("colsTargetDelete", parameters.get("colsTargetDelete").get());
                  if(convertedParameters.get("colsTargetDelete") !=null && !(convertedParameters.get("colsTargetDelete") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","colsTargetDelete", "String", parameters.get("colsTargetDelete").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("colsTargetDelete") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","colsTargetDelete"));
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","colOperation"));
            }
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","group"));
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
      command.action((String)convertedParameters.get("group"),(String)convertedParameters.get("rowOperation"),(String)convertedParameters.get("rowsTargetInsert"),(String)convertedParameters.get("rowsTargetDelete"),(String)convertedParameters.get("colOperation"),(String)convertedParameters.get("colsTargetInsert"),(String)convertedParameters.get("colsTargetDelete"),(WorkbookSession)convertedParameters.get("session"));Optional<Value> result = Optional.empty();
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
