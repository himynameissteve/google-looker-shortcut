const cc = DataStudioApp.createCommunityConnector();

function isAdminUser() {
  return true;
}

/**
 * Mandatory function of Google Looker.
 * Configures the authentication type of the API connection (Shortcut).
 *
 * @returns Authentication type for Shortcut.
 */
function getAuthType() {
  const AuthTypes = cc.AuthType;

  return cc
    .newAuthTypeResponse()
    .setAuthType(AuthTypes.KEY)
    .setHelpUrl("https://developer.shortcut.com/api/rest/v3#Authentication")
    .build();
}

/**
 * Mandatory function of Google Looker for certain authentication types.
 * Checks if the currently set authentication configuration is still valid.
 * Invalid keys trigger a reconfiguration in Google Looker.
 *
 * @returns true if Shortcut API key is valid.
 */
function isAuthValid() {
  const apiKey = PropertiesService.getUserProperties().getProperty("apiKey");
  return apiKey == null || apiKey == "" ? false : checkApiKeyValidity(apiKey);
}

/**
 * Check if the given Shortcut API key is able to retrieve data from Shortcut
 * by calling a random (categories) endpoint.
 * Hopefully they won't change the name or URL of the endpoint.
 *
 * @param apiKey for Shortcut.
 * @returns true if the Shortcut API call produces a HTTP-200 result.
 */
function checkApiKeyValidity(apiKey) {
  const url = "https://api.app.shortcut.com/api/v3/categories";
  const options = {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
      "Shortcut-Token": apiKey,
    },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  return response.getResponseCode() == 200;
}

/**
 * Mandatory function of Google Looker.
 * Deletes the set Shortcut API key from internal storage.
 */
function resetAuth() {
  PropertiesService.getUserProperties().deleteProperty("apiKey");
}

/**
 * Mandatory function of Google Looker.
 * Sets the given Shortcut API key as a user property for further use in Shortcut api calls.
 *
 * @param request set by Google Looker. Holds the given Shortcut api key.
 * @returns error code object with predefined status codes that Google Looker can process.
 */
function setCredentials(request) {
  const apiKey = request.key;

  if (!checkApiKeyValidity(apiKey)) {
    return {
      errorCode: "INVALID_CREDENTIALS",
    };
  }

  let userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty("apiKey", apiKey);
  return {
    errorCode: "NONE",
  };
}

/**
 * Mandatory function by Google Looker.
 * Configures the prompted fields during set up of the plugin/add-on.
 *
 * @param request given by Google Looker.
 * @returns config object.
 */
function getConfig(request) {
  const config = cc.getConfig();

  config
    .newInfo()
    .setId("instructions")
    .setText(
      "The Shortcut API paginates its results. That is limited to a certain degree and results in an error after that. If data retrieval results in an error, lower the data range for the request to receive fewer results."
    );

  config
    .newTextInput()
    .setId("project")
    .setName("Enter a single Shortcut project name")
    .setHelpText("This is a project you configured in Shortcut.");

  config.setDateRangeRequired(true);

  return config.build();
}

/**
 * Mandatory function of Google Looker.
 *
 * Configures the fields you want to pull and map from Shortcut and their types.
 * This is what Google Looker later identifies as dimensions and metrics for their visualizations.
 *
 * @param request given by Google Looker.
 * @returns field configuration.
 */
function getFields(request) {
  const cc = DataStudioApp.createCommunityConnector();
  let fields = cc.getFields();
  const types = cc.FieldType;
  const aggregations = cc.AggregationType;

  fields.newDimension().setId("completed").setType(types.YEAR_MONTH_DAY);

  fields.newDimension().setId("created").setType(types.YEAR_MONTH_DAY);

  fields.newDimension().setId("teams").setType(types.TEXT);

  fields.newDimension().setId("storyType").setType(types.TEXT);

  fields
    .newMetric()
    .setId("count")
    .setName("Number of stories")
    .setType(types.NUMBER);

  return fields;
}

/**
 * Mandatory function of Google Looker.
 *
 * Defines the database schema based on the configured fields.
 *
 * @param request given by Google Looker.
 * @returns schema configuration.
 */
function getSchema(request) {
  const fields = getFields(request).build();
  return { schema: fields };
}

/**
 * Maps Shortcut group ids to readable team names.
 *
 * @param apiKey for Shortcut API calls.
 * @returns Map of group id to team name mappings.
 */
function getGroupsFromShortcut(apiKey) {
  const url = "https://api.app.shortcut.com/api/v3/groups";
  const options = {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
      "Shortcut-Token": apiKey,
    },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() == 200) {
    const parsedResponse = JSON.parse(response);
    const mapping = new Map();
    parsedResponse.forEach((group) => {
      mapping.set(group.id, group.name);
    });

    return mapping;
  } else {
    DataStudioApp.createCommunityConnector()
      .newUserError()
      .setDebugText(
        "Error fetching data from Shortcut API. Exception details: " + response
      )
      .setText(
        "Error fetching data from Shortcut API. Exception details: " + response
      )
      .throwException();
  }
}

function responseToRows(requestedFields, response, projectName, groupMapping) {
  // Transform parsed data and filter for requested fields
  return response.map(function (story) {
    const row = [];
    requestedFields.asArray().forEach(function (field) {
      switch (field.getId()) {
        case "completed":
          if (story.completed_at == null) {
            return row.push("");
          }
          return row.push(
            story.completed_at.substring(0, 10).replace(/-/g, "")
          );
        case "created":
          if (story.created_at == null) {
            return row.push("");
          }
          return row.push(story.created_at.substring(0, 10).replace(/-/g, ""));
        case "teams":
          if (story.group_id == null) {
            return row.push("");
          }
          return row.push(groupMapping.get(story.group_id));
        case "storyType":
          if (story.story_type == null) {
            return row.push("");
          }
          return row.push(story.story_type);
        case "count":
          return row.push(1);
        default:
          return row.push("");
      }
    });
    return { values: row };
  });
}

/**
 * Mandatory function of Google Looker.
 *
 * Get the Shortcut data and maps it to Google Looker's schema and field configuration.
 *
 * @param request provided by Google Looker.
 * @returns Shortcut data in Google Looker configured representation.
 */
function getData(request) {
  const requestedFieldIds = request.fields.map(function (field) {
    return field.name;
  });
  const requestedFields = getFields().forIds(requestedFieldIds);
  const apiKey = PropertiesService.getUserProperties().getProperty("apiKey");
  const project = request.configParams.project;
  const startDate = request.dateRange.startDate;
  const endDate = request.dateRange.endDate;

  const baseUrl = "https://api.app.shortcut.com";
  let path =
    "/api/v3/search/stories?detail=full&page_size=25&query=project%3A" +
    project +
    "%20AND%20completed%3A" +
    encodeURIComponent(startDate) +
    "%2E%2E" +
    encodeURIComponent(endDate);
  const options = {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
      "Shortcut-Token": apiKey,
    },
    muteHttpExceptions: true,
  };
  const groupMapping = getGroupsFromShortcut(apiKey);
  let next = "";

  let allRows = [];

  do {
    let response = UrlFetchApp.fetch(baseUrl + path, options);

    if (response.getResponseCode() == 200) {
      Logger.log("Current path: " + path);
      const parsedResponse = JSON.parse(response);
      const stories = parsedResponse.data;
      next = parsedResponse.next;
      path = next;
      allRows = allRows.concat(
        responseToRows(
          requestedFields,
          stories,
          request.configParams.project,
          groupMapping
        )
      );
    } else {
      DataStudioApp.createCommunityConnector()
        .newUserError()
        .setDebugText(
          "Error fetching data from Shortcut API. Exception details: " +
            response
        )
        .setText(
          "Error fetching data from Shortcut API. Exception details: " +
            response
        )
        .throwException();
    }
  } while (next);

  return {
    schema: requestedFields.build(),
    rows: allRows,
  };
}
