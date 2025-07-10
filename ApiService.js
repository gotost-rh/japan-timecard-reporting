/**
 * ApiService - Handles all API interactions with PSA (Professional Services Automation)
 * Provides standardized methods for querying data with proper error handling and pagination
 */

/**
 * Configuration for API service
 */
const API_CONFIG = {
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 1000,
  BATCH_SIZE: 2000
};

/**
 * Execute a SOQL query with automatic pagination handling
 * @param {string} queryString - The SOQL query string
 * @returns {Array} Array of all records from the query
 * @throws {Error} If query fails after retries
 */
function executeQuery(queryString) {
  console.log(`Executing query: ${queryString}`);
  
  try {
    let results = executeQueryWithRetry_(queryString);
    const allRecords = [];

    // Add initial batch of records
    if (results && results.records) {
      allRecords.push(...results.records);
    }
    
    // Handle pagination if more records exist
    if (results && results.done === false) {
      console.log(`Query returned ${allRecords.length} records initially, fetching remaining...`);
      
      do {
        results = executeQueryMoreWithRetry_(results.nextRecordsUrl);
        if (results && results.records) {
          allRecords.push(...results.records);
        }
      } while (results && results.done === false);
    }

    console.log(`Query completed. Total records retrieved: ${allRecords.length}`);
    return allRecords;
    
  } catch (error) {
    console.error(`Query execution failed: ${error.message}`);
    throw new Error(`Failed to execute query: ${error.message}`);
  }
}

/**
 * Execute a SOQL query for timecards with date filtering
 * @param {string} opportunityNumber - The opportunity number to filter by
 * @param {Date} startDate - Start date for timecard filtering
 * @param {Date} endDate - End date for timecard filtering
 * @returns {Array} Array of timecard records
 */
function queryTimecards(opportunityNumber, startDate, endDate) {
  const queryString = buildTimecardQuery_(opportunityNumber, startDate, endDate);
  return executeQuery(queryString);
}

/**
 * Execute a SOQL query for projects
 * @param {string} opportunityNumber - The opportunity number to filter by
 * @returns {Array} Array of project records
 */
function queryProjects(opportunityNumber) {
  const queryString = buildProjectQuery_(opportunityNumber);
  return executeQuery(queryString);
}

/**
 * Build SOQL query for timecards
 * @private
 * @param {string} opportunityNumber - The opportunity number
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @returns {string} SOQL query string
 */
function buildTimecardQuery_(opportunityNumber, startDate, endDate) {
  const queryFields = TIMECARD_FIELDS.join(",");
  const baseFilter = TIMECARD_FILTER.join("+AND+");
  const formattedStartDate = formatDateForQuery_(startDate);
  const formattedEndDate = formatDateForQuery_(endDate);
  
  const queryFilter = `${baseFilter}+AND+pse__Project__r.pse__Opportunity__r.OpportunityNumber__c%3D'${opportunityNumber}'` +
                     `+AND+pse__Start_Date__c%3E%3D${formattedStartDate}` +
                     `+AND+pse__Start_Date__c%3C%3D${formattedEndDate}`;
  
  return `SELECT+${queryFields}+FROM+${TIMECARD_API_NAME}+WHERE+${queryFilter}+ORDER+BY+pse__Start_Date__c`;
}

/**
 * Build SOQL query for projects
 * @private
 * @param {string} opportunityNumber - The opportunity number
 * @returns {string} SOQL query string
 */
function buildProjectQuery_(opportunityNumber) {
  const queryFields = PROJECT_FIELDS.join(",");
  const queryFilter = `pse__Opportunity__r.OpportunityNumber__c%3D'${opportunityNumber}'`;
  
  return `SELECT+${queryFields}+FROM+${PROJECT_API_NAME}+WHERE+${queryFilter}`;
}

/**
 * Format date for SOQL query
 * @private
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string for SOQL
 */
function formatDateForQuery_(date) {
  if (!date) return null;
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

/**
 * Execute RHCCRestApi.query with retry logic
 * @private
 * @param {string} queryString - The query string
 * @returns {Object} API response
 * @throws {Error} If all retries fail
 */
function executeQueryWithRetry_(queryString) {
  let lastError;
  
  for (let attempt = 1; attempt <= API_CONFIG.MAX_RETRIES; attempt++) {
    try {
      return RHCCRestApi.query(queryString);
    } catch (error) {
      lastError = error;
      console.warn(`Query attempt ${attempt} failed: ${error.message}`);
      
      if (attempt < API_CONFIG.MAX_RETRIES) {
        console.log(`Retrying in ${API_CONFIG.RETRY_DELAY_MS}ms...`);
        Utilities.sleep(API_CONFIG.RETRY_DELAY_MS);
      }
    }
  }
  
  throw new Error(`Query failed after ${API_CONFIG.MAX_RETRIES} attempts. Last error: ${lastError.message}`);
}

/**
 * Execute RHCCRestApi.queryMore with retry logic
 * @private
 * @param {string} nextRecordsUrl - The next records URL
 * @returns {Object} API response
 * @throws {Error} If all retries fail
 */
function executeQueryMoreWithRetry_(nextRecordsUrl) {
  let lastError;
  
  for (let attempt = 1; attempt <= API_CONFIG.MAX_RETRIES; attempt++) {
    try {
      return RHCCRestApi.queryMore(nextRecordsUrl);
    } catch (error) {
      lastError = error;
      console.warn(`QueryMore attempt ${attempt} failed: ${error.message}`);
      
      if (attempt < API_CONFIG.MAX_RETRIES) {
        console.log(`Retrying in ${API_CONFIG.RETRY_DELAY_MS}ms...`);
        Utilities.sleep(API_CONFIG.RETRY_DELAY_MS);
      }
    }
  }
  
  throw new Error(`QueryMore failed after ${API_CONFIG.MAX_RETRIES} attempts. Last error: ${lastError.message}`);
}

/**
 * Utility function to format date (keeping backward compatibility)
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string
 */
function formatDate(date) {
  return formatDateForQuery_(date);
} 