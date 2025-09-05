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

public final class SortWorksheetOrTableCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(SortWorksheetOrTableCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    SortWorksheetOrTable command = new SortWorksheetOrTable();
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
    if(convertedParameters.get("mode") != null) {
      switch((String)convertedParameters.get("mode")) {
        case "TABLE" : {
          if(parameters.containsKey("tableName") && parameters.get("tableName") != null && parameters.get("tableName").get() != null) {
            convertedParameters.put("tableName", parameters.get("tableName").get());
            if(convertedParameters.get("tableName") !=null && !(convertedParameters.get("tableName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableName", "String", parameters.get("tableName").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("tableName") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","tableName"));
          }

          if(parameters.containsKey("tableColSelector") && parameters.get("tableColSelector") != null && parameters.get("tableColSelector").get() != null) {
            convertedParameters.put("tableColSelector", parameters.get("tableColSelector").get());
            if(convertedParameters.get("tableColSelector") !=null && !(convertedParameters.get("tableColSelector") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableColSelector", "String", parameters.get("tableColSelector").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("tableColSelector") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","tableColSelector"));
          }
          if(convertedParameters.get("tableColSelector") != null) {
            switch((String)convertedParameters.get("tableColSelector")) {
              case "BY_NAME" : {
                if(parameters.containsKey("tableColumnName") && parameters.get("tableColumnName") != null && parameters.get("tableColumnName").get() != null) {
                  convertedParameters.put("tableColumnName", parameters.get("tableColumnName").get());
                  if(convertedParameters.get("tableColumnName") !=null && !(convertedParameters.get("tableColumnName") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableColumnName", "String", parameters.get("tableColumnName").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("tableColumnName") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","tableColumnName"));
                }


              } break;
              case "BY_INDEX" : {
                if(parameters.containsKey("tableColumnIndexOneBased") && parameters.get("tableColumnIndexOneBased") != null && parameters.get("tableColumnIndexOneBased").get() != null) {
                  convertedParameters.put("tableColumnIndexOneBased", parameters.get("tableColumnIndexOneBased").get());
                  if(convertedParameters.get("tableColumnIndexOneBased") !=null && !(convertedParameters.get("tableColumnIndexOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableColumnIndexOneBased", "Double", parameters.get("tableColumnIndexOneBased").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("tableColumnIndexOneBased") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","tableColumnIndexOneBased"));
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","tableColSelector"));
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
          if(convertedParameters.get("sheetName") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sheetName"));
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


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","wsRangeMode"));
            }
          }

          if(parameters.containsKey("wsColSelector") && parameters.get("wsColSelector") != null && parameters.get("wsColSelector").get() != null) {
            convertedParameters.put("wsColSelector", parameters.get("wsColSelector").get());
            if(convertedParameters.get("wsColSelector") !=null && !(convertedParameters.get("wsColSelector") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsColSelector", "String", parameters.get("wsColSelector").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("wsColSelector") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsColSelector"));
          }
          if(convertedParameters.get("wsColSelector") != null) {
            switch((String)convertedParameters.get("wsColSelector")) {
              case "BY_NAME" : {
                if(parameters.containsKey("wsColumnName") && parameters.get("wsColumnName") != null && parameters.get("wsColumnName").get() != null) {
                  convertedParameters.put("wsColumnName", parameters.get("wsColumnName").get());
                  if(convertedParameters.get("wsColumnName") !=null && !(convertedParameters.get("wsColumnName") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsColumnName", "String", parameters.get("wsColumnName").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("wsColumnName") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsColumnName"));
                }


              } break;
              case "BY_INDEX" : {
                if(parameters.containsKey("wsColumnIndexOneBased") && parameters.get("wsColumnIndexOneBased") != null && parameters.get("wsColumnIndexOneBased").get() != null) {
                  convertedParameters.put("wsColumnIndexOneBased", parameters.get("wsColumnIndexOneBased").get());
                  if(convertedParameters.get("wsColumnIndexOneBased") !=null && !(convertedParameters.get("wsColumnIndexOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsColumnIndexOneBased", "Double", parameters.get("wsColumnIndexOneBased").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("wsColumnIndexOneBased") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsColumnIndexOneBased"));
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","wsColSelector"));
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


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","mode"));
      }
    }

    if(parameters.containsKey("sortType") && parameters.get("sortType") != null && parameters.get("sortType").get() != null) {
      convertedParameters.put("sortType", parameters.get("sortType").get());
      if(convertedParameters.get("sortType") !=null && !(convertedParameters.get("sortType") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sortType", "String", parameters.get("sortType").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("sortType") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sortType"));
    }
    if(convertedParameters.get("sortType") != null) {
      switch((String)convertedParameters.get("sortType")) {
        case "NUMBER" : {
          if(parameters.containsKey("numberOrder") && parameters.get("numberOrder") != null && parameters.get("numberOrder").get() != null) {
            convertedParameters.put("numberOrder", parameters.get("numberOrder").get());
            if(convertedParameters.get("numberOrder") !=null && !(convertedParameters.get("numberOrder") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","numberOrder", "String", parameters.get("numberOrder").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("numberOrder") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","numberOrder"));
          }
          if(convertedParameters.get("numberOrder") != null) {
            switch((String)convertedParameters.get("numberOrder")) {
              case "ASC" : {

              } break;
              case "DESC" : {

              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","numberOrder"));
            }
          }


        } break;
        case "TEXT" : {
          if(parameters.containsKey("textOrder") && parameters.get("textOrder") != null && parameters.get("textOrder").get() != null) {
            convertedParameters.put("textOrder", parameters.get("textOrder").get());
            if(convertedParameters.get("textOrder") !=null && !(convertedParameters.get("textOrder") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","textOrder", "String", parameters.get("textOrder").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("textOrder") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","textOrder"));
          }
          if(convertedParameters.get("textOrder") != null) {
            switch((String)convertedParameters.get("textOrder")) {
              case "ASC" : {

              } break;
              case "DESC" : {

              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","textOrder"));
            }
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","sortType"));
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
      command.action((String)convertedParameters.get("mode"),(String)convertedParameters.get("tableName"),(String)convertedParameters.get("tableColSelector"),(String)convertedParameters.get("tableColumnName"),(Double)convertedParameters.get("tableColumnIndexOneBased"),(String)convertedParameters.get("sheetName"),(String)convertedParameters.get("wsRangeMode"),(String)convertedParameters.get("wsRangeA1"),(String)convertedParameters.get("wsColSelector"),(String)convertedParameters.get("wsColumnName"),(Double)convertedParameters.get("wsColumnIndexOneBased"),(Boolean)convertedParameters.get("wsHasHeader"),(String)convertedParameters.get("sortType"),(String)convertedParameters.get("numberOrder"),(String)convertedParameters.get("textOrder"),(WorkbookSession)convertedParameters.get("session"));Optional<Value> result = Optional.empty();
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
