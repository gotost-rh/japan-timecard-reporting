/**
 * TestGPSMapping - Unit tests for GPS code mapping functionality
 * Run these tests by calling runAllGPSTests() from the script editor
 */

/**
 * Test helper function to assert equality
 * @param {*} actual - Actual value
 * @param {*} expected - Expected value
 * @param {string} testName - Name of the test
 */
function assertEqual(actual, expected, testName) {
  if (actual === expected) {
    console.log(`✅ PASS: ${testName}`);
    return true;
  } else {
    console.error(`❌ FAIL: ${testName}`);
    console.error(`   Expected: "${expected}"`);
    console.error(`   Actual: "${actual}"`);
    return false;
  }
}

/**
 * Test exact matches from GPS_CODE_MAPPING
 */
function testExactMatches() {
  console.log("\n🔍 Testing exact matches...");
  
  const tests = [
    { input: "project manager", expected: "GPS-PJM" },
    { input: "principal consultant", expected: "GPS-PRCON" },
    { input: "senior consultant", expected: "GPS-SC" },
    { input: "consultant", expected: "GPS-C" },
    { input: "principal architect", expected: "GPS-PA" },
    { input: "senior architect", expected: "GPS-SA" },
    { input: "architect", expected: "GPS-A" }
  ];
  
  let passed = 0;
  for (const test of tests) {
    if (assertEqual(mapToGPSCode(test.input), test.expected, `Exact match: "${test.input}"`)) {
      passed++;
    }
  }
  
  console.log(`Exact matches: ${passed}/${tests.length} tests passed`);
  return passed === tests.length;
}

/**
 * Test extracting existing GPS codes from milestone names
 */
function testExistingGPSCodes() {
  console.log("\n🔍 Testing existing GPS code extraction...");
  
  const tests = [
    { input: "GPS-PJM for TMNS_PO Essential_CY25Q2 - Milestone", expected: "GPS-PJM" },
    { input: "GPS-SC - Technical Implementation Phase", expected: "GPS-SC" },
    { input: "Project GPS-A Architecture Review", expected: "GPS-A" },
    { input: "GPS-PRCON Consulting Services", expected: "GPS-PRCON" },
    { input: "Task GPS-PA Planning Activities", expected: "GPS-PA" },
    { input: "gps-pjm for lower case test", expected: "GPS-PJM" }, // Case insensitive
    { input: "Multiple GPS-SC and GPS-A codes", expected: "GPS-SC" }, // Should return first match
  ];
  
  let passed = 0;
  for (const test of tests) {
    if (assertEqual(mapToGPSCode(test.input), test.expected, `Extract GPS code: "${test.input}"`)) {
      passed++;
    }
  }
  
  console.log(`GPS code extraction: ${passed}/${tests.length} tests passed`);
  return passed === tests.length;
}

/**
 * Test unmapped roles return original name
 */
function testUnmappedRoles() {
  console.log("\n🔍 Testing unmapped roles...");
  
  const tests = [
    { input: "Custom Role Name", expected: "Custom Role Name" },
    { input: "Random Job Title", expected: "Random Job Title" },
    { input: "Specialized Engineer", expected: "Specialized Engineer" },
    { input: "Business Analyst Level 2", expected: "Business Analyst Level 2" },
    { input: "Data Scientist", expected: "Data Scientist" }
  ];
  
  let passed = 0;
  for (const test of tests) {
    if (assertEqual(mapToGPSCode(test.input), test.expected, `Unmapped role: "${test.input}"`)) {
      passed++;
    }
  }
  
  console.log(`Unmapped roles: ${passed}/${tests.length} tests passed`);
  return passed === tests.length;
}

/**
 * Test case sensitivity
 */
function testCaseSensitivity() {
  console.log("\n🔍 Testing case sensitivity...");
  
  const tests = [
    { input: "PROJECT MANAGER", expected: "GPS-PJM" },
    { input: "Principal Consultant", expected: "GPS-PRCON" },
    { input: "SENIOR ARCHITECT", expected: "GPS-SA" },
    { input: "Architect", expected: "GPS-A" }
  ];
  
  let passed = 0;
  for (const test of tests) {
    if (assertEqual(mapToGPSCode(test.input), test.expected, `Case insensitive: "${test.input}"`)) {
      passed++;
    }
  }
  
  console.log(`Case sensitivity: ${passed}/${tests.length} tests passed`);
  return passed === tests.length;
}

/**
 * Test partial matches (substring matching)
 */
function testPartialMatches() {
  console.log("\n🔍 Testing partial matches...");
  
  const tests = [
    { input: "Senior Consultant Level 3", expected: "GPS-SC" },
    { input: "Principal Architect Team Lead", expected: "GPS-PA" },
    { input: "Project Manager - Implementation", expected: "GPS-PJM" },
    { input: "Technical Architect Role", expected: "GPS-A" },
    { input: "GPS-PJM for TMNS_PO Essential_CY25Q2 - Milestone", expected: "GPS-PJM" },
  ];
  
  let passed = 0;
  for (const test of tests) {
    if (assertEqual(mapToGPSCode(test.input), test.expected, `Partial match: "${test.input}"`)) {
      passed++;
    }
  }
  
  console.log(`Partial matches: ${passed}/${tests.length} tests passed`);
  return passed === tests.length;
}

/**
 * Test edge cases and invalid inputs
 */
function testEdgeCases() {
  console.log("\n🔍 Testing edge cases...");
  
  const tests = [
    { input: "", expected: "" },
    { input: null, expected: null },
    { input: undefined, expected: undefined },
    { input: "   ", expected: "   " }, // whitespace
    { input: "GPS- incomplete", expected: "GPS- incomplete" }, // Invalid GPS pattern
    { input: "GPS-", expected: "GPS-" }, // Incomplete GPS code
    { input: "No GPS code here", expected: "No GPS code here" } // No GPS pattern
  ];
  
  let passed = 0;
  for (const test of tests) {
    if (assertEqual(mapToGPSCode(test.input), test.expected, `Edge case: "${test.input}"`)) {
      passed++;
    }
  }
  
  console.log(`Edge cases: ${passed}/${tests.length} tests passed`);
  return passed === tests.length;
}

/**
 * Test whitespace handling
 */
function testWhitespaceHandling() {
  console.log("\n🔍 Testing whitespace handling...");
  
  const tests = [
    { input: "  project manager  ", expected: "GPS-PJM" },
    { input: "\tconsultant\n", expected: "GPS-C" },
    { input: " principal   consultant ", expected: "GPS-PRCON" }
  ];
  
  let passed = 0;
  for (const test of tests) {
    if (assertEqual(mapToGPSCode(test.input), test.expected, `Whitespace: "${test.input}"`)) {
      passed++;
    }
  }
  
  console.log(`Whitespace handling: ${passed}/${tests.length} tests passed`);
  return passed === tests.length;
}

/**
 * Test priority order (more specific matches should win)
 */
function testPriorityOrder() {
  console.log("\n🔍 Testing priority order...");
  
  const tests = [
    // "principal consultant" should match before "consultant"
    { input: "principal consultant", expected: "GPS-PRCON" },
    // "senior architect" should match before "architect"
    { input: "senior architect", expected: "GPS-SA" },
    // "project manager" should match before any generic matches
    { input: "project manager", expected: "GPS-PJM" }
  ];
  
  let passed = 0;
  for (const test of tests) {
    if (assertEqual(mapToGPSCode(test.input), test.expected, `Priority: "${test.input}"`)) {
      passed++;
    }
  }
  
  console.log(`Priority order: ${passed}/${tests.length} tests passed`);
  return passed === tests.length;
}

/**
 * Performance test - test with many inputs
 */
function testPerformance() {
  console.log("\n🔍 Testing performance...");
  
  const startTime = Date.now();
  const testInputs = [
    "project manager", "consultant", "architect", "principal consultant",
    "senior architect", "random role", "another role", "test role"
  ];
  
  // Run 1000 iterations
  for (let i = 0; i < 1000; i++) {
    for (const input of testInputs) {
      mapToGPSCode(input);
    }
  }
  
  const endTime = Date.now();
  const duration = endTime - startTime;
  
  console.log(`✅ Performance test: ${8000} calls completed in ${duration}ms`);
  console.log(`   Average: ${(duration / 8000).toFixed(3)}ms per call`);
  
  return duration < 5000; // Should complete within 5 seconds
}

/**
 * Run all GPS mapping tests
 * Call this function from the Apps Script editor to run all tests
 */
function runAllGPSTests() {
  console.log("🧪 Starting GPS Code Mapping Tests...");
  console.log("=" * 50);
  
  const testResults = [];
  
  // Run all test suites
  testResults.push(testExactMatches());
  testResults.push(testExistingGPSCodes()); // New test
  testResults.push(testUnmappedRoles()); // New test
  testResults.push(testCaseSensitivity());
  testResults.push(testPartialMatches());
  testResults.push(testEdgeCases());
  testResults.push(testWhitespaceHandling());
  testResults.push(testPriorityOrder());
  testResults.push(testPerformance());
  
  // Summary
  const passedSuites = testResults.filter(result => result === true).length;
  const totalSuites = testResults.length;
  
  console.log("\n" + "=" * 50);
  console.log(`🎯 Test Summary: ${passedSuites}/${totalSuites} test suites passed`);
  
  if (passedSuites === totalSuites) {
    console.log("🎉 All tests passed!");
  } else {
    console.log("⚠️  Some tests failed. Please review the output above.");
  }
  
  return passedSuites === totalSuites;
}

/**
 * Quick test function for manual testing
 * @param {string} input - Input to test
 */
function testSingleInput(input) {
  console.log(`Testing: "${input}"`);
  const result = mapToGPSCode(input);
  console.log(`Result: "${result}"`);
  return result;
}

/**
 * Test with actual milestone names from your data
 * Add real milestone names from your system here
 */
function testWithRealData() {
  console.log("\n🔍 Testing with sample real data...");
  
  // Replace these with actual milestone names from your system
  const realMilestoneNames = [
    "GPS-PJM for TMNS_PO Essential_CY25Q2 - Milestone",
    "GPS-SC - Technical Implementation Phase",
    "Project Management - Phase 1",
    "Senior Consultant - Technical Review",
    "Principal Architect - Solution Design",
    "Implementation Support",
    "Quality Assurance Testing"
  ];
  
  for (const milestoneName of realMilestoneNames) {
    const result = mapToGPSCode(milestoneName);
    console.log(`"${milestoneName}" → "${result}"`);
  }
}

/**
 * Interactive test function for debugging
 * Run this with different inputs to see detailed mapping logic
 */
function debugGPSMapping(input) {
  console.log(`\n🔍 Debugging GPS mapping for: "${input}"`);
  
  if (!input || typeof input !== 'string') {
    console.log("❌ Invalid input - returning original");
    return input;
  }
  
  const lowerName = input.toLowerCase().trim();
  console.log(`📝 Processed input: "${lowerName}"`);
  
  // Check exact matches
  if (GPS_CODE_MAPPING[lowerName]) {
    console.log(`✅ Exact match found: ${GPS_CODE_MAPPING[lowerName]}`);
    return GPS_CODE_MAPPING[lowerName];
  }
  
  // Check partial matches
  console.log("🔍 Checking partial matches...");
  for (const [key, value] of Object.entries(GPS_CODE_MAPPING)) {
    if (lowerName.includes(key)) {
      console.log(`✅ Partial match found: "${key}" → ${value}`);
      return value;
    }
  }
  
  console.log("❌ No mapping found - returning original");
  return input;
} 