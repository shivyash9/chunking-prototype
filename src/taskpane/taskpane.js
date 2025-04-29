/* global Word console */

// Function to read all paragraphs and assign unique IDs using content controls
export async function processParagraphs() {
  const startTime = Date.now();
  const paragraphData = [];
  const processedTables = new Set(); 
  const processedContentControls = new Set(); 
  const tableIdToTagMap = new Map();
  
  try {
    return await Word.run(async (context) => {
      // PHASE 1: Load all content controls, tables, and paragraphs in the document
      console.log("Phase 1: Loading document elements");
      
      // Get existing content controls
      const existingContentControls = context.document.body.contentControls;
      existingContentControls.load("items");
      
      // Load tables
      const allTables = context.document.body.tables;
      allTables.load("items");
      
      // Load paragraphs
      const allParagraphs = context.document.body.paragraphs;
      allParagraphs.load("items");
      
      // Execute sync for initial loading
      await context.sync();
      console.log(`Loaded ${existingContentControls.items ? existingContentControls.items.length : 0} content controls`);
      console.log(`Loaded ${allTables.items ? allTables.items.length : 0} tables`);
      console.log(`Loaded ${allParagraphs.items ? allParagraphs.items.length : 0} paragraphs`);
      
      // PHASE 2: Process existing content controls
      console.log("Phase 2: Processing existing content controls");
      
      if (existingContentControls.items && existingContentControls.items.length > 0) {
        // Load all tags in a single batch
        for (let i = 0; i < existingContentControls.items.length; i++) {
          if (existingContentControls.items[i]) {
            existingContentControls.items[i].load("tag");
          }
        }
        await context.sync();
        
        // First collect all content controls with our tags
        for (let i = 0; i < existingContentControls.items.length; i++) {
          const cc = existingContentControls.items[i];
          if (cc && cc.tag) {
            if (cc.tag.startsWith('para-') || cc.tag.startsWith('table-')) {
              processedContentControls.add(cc.tag);
            }
          }
        }
        
        // Now handle table content controls separately to build table ID mapping
        const tableContentControls = [];
        for (let i = 0; i < existingContentControls.items.length; i++) {
          const cc = existingContentControls.items[i];
          if (cc && cc.tag && cc.tag.startsWith('table-')) {
            tableContentControls.push(cc);
          }
        }
        
        if (tableContentControls.length > 0) {
          console.log(`Found ${tableContentControls.length} table content controls to process`);
          
          // Process in smaller batches to avoid timeouts
          const batchSize = 5;
          for (let i = 0; i < tableContentControls.length; i += batchSize) {
            try {
              const batch = tableContentControls.slice(i, i + batchSize);
              
              // For each content control, get its range and load tables
              for (const cc of batch) {
                if (!cc) continue;
                const range = cc.getRange();
                if (!range) continue;
                const tables = range.tables;
                tables.load("items");
              }
              await context.sync();
              
              // Now process each content control in this batch
              for (const cc of batch) {
                try {
                  if (!cc) continue;
                  const range = cc.getRange();
                  if (!range) continue;
                  const tables = range.tables;
                  
                  // We need to load tables.items again here because access to items in a different context
                  tables.load("items");
                  await context.sync();
                  
                  if (tables && tables.items && tables.items.length > 0) {
                    // Load table IDs
                    for (let j = 0; j < tables.items.length; j++) {
                      if (tables.items[j]) {
                        tables.items[j].load("id");
                      }
                    }
                    await context.sync();
                    
                    // Map first table's ID to this content control tag
                    if (tables.items[0] && tables.items[0].id) {
                      const tableId = tables.items[0].id;
                      tableIdToTagMap.set(tableId, cc.tag);
                    }
                  }
                } catch (e) {
                  console.error("Error mapping table content control to table ID:", e);
                }
              }
            } catch (e) {
              console.error("Error processing table content control batch:", e);
            }
          }
        }
      }
      
      // PHASE 3: Process tables that don't have content controls yet
      console.log("Phase 3: Processing tables");
      
      if (allTables.items && allTables.items.length > 0) {
        console.log(`Processing ${allTables.items.length} tables`);
        
        // Load all table IDs first
        for (let i = 0; i < allTables.items.length; i++) {
          const table = allTables.items[i];
          if (table) {
            table.load("id");
          }
        }
        await context.sync();
        
        // Process tables in batches
        const tableBatchSize = 5;
        for (let batchStart = 0; batchStart < allTables.items.length; batchStart += tableBatchSize) {
          const batchEnd = Math.min(batchStart + tableBatchSize, allTables.items.length);
          console.log(`Processing table batch ${batchStart}-${batchEnd-1}`);
          
          // Check for existing content controls in this batch
          for (let t = batchStart; t < batchEnd; t++) {
            const table = allTables.items[t];
            if (!table || processedTables.has(table.id)) continue;
            
            // Check if already mapped from earlier
            if (!tableIdToTagMap.has(table.id)) {
              try {
                // Check for content controls on this table
                const tableContentControls = table.contentControls;
                if (tableContentControls) {
                  tableContentControls.load("items");
                }
              } catch (e) {
                console.error("Error loading table content controls:", e);
              }
            }
          }
          await context.sync();
          
          // Now process each table in the batch
          for (let t = batchStart; t < batchEnd; t++) {
            try {
              const table = allTables.items[t];
              if (!table || processedTables.has(table.id)) continue;
              
              // Skip if already mapped to a content control
              if (tableIdToTagMap.has(table.id)) {
                const uniqueId = tableIdToTagMap.get(table.id);
                processedTables.add(table.id);
                processedContentControls.add(uniqueId);
                continue;
              }
              
              // Check existing content controls on this table
              const tableContentControls = table.contentControls;
              if (tableContentControls && tableContentControls.items && tableContentControls.items.length > 0) {
                // Load tags on these content controls
                for (let c = 0; c < tableContentControls.items.length; c++) {
                  if (tableContentControls.items[c]) {
                    tableContentControls.items[c].load("tag");
                  }
                }
                await context.sync();
                
                // Check if any has our expected tag format
                let hasOurContentControl = false;
                for (let c = 0; c < tableContentControls.items.length; c++) {
                  const cc = tableContentControls.items[c];
                  if (cc && cc.tag && cc.tag.startsWith("table-")) {
                    tableIdToTagMap.set(table.id, cc.tag);
                    processedContentControls.add(cc.tag);
                    hasOurContentControl = true;
                    break;
                  }
                }
                
                if (hasOurContentControl) {
                  processedTables.add(table.id);
                  continue;
                }
              }
              
              // If we get here, we need to add a new content control
              const uniqueId = `table-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
              try {
                const tableCC = table.insertContentControl();
                tableCC.tag = uniqueId;
                tableCC.title = "Table " + (t + 1);
                tableIdToTagMap.set(table.id, uniqueId);
                processedTables.add(table.id);
                processedContentControls.add(uniqueId);
              } catch (e) {
                console.error("Error inserting table content control:", e);
              }
            } catch (e) {
              console.error("Error processing table:", e);
            }
          }
          
          // Sync at the end of each batch
          await context.sync();
        }
      }
      
      // PHASE 4: Process paragraphs that aren't already in content controls or tables
      console.log("Phase 4: Processing paragraphs");
      
      if (allParagraphs.items && allParagraphs.items.length > 0) {
        console.log(`Processing ${allParagraphs.items.length} paragraphs`);
        
        // Process paragraphs in smaller batches
        const paragraphBatchSize = 10;
        for (let batchStart = 0; batchStart < allParagraphs.items.length; batchStart += paragraphBatchSize) {
          const batchEnd = Math.min(batchStart + paragraphBatchSize, allParagraphs.items.length);
          console.log(`Processing paragraph batch ${batchStart}-${batchEnd-1}`);
          
          // Load text and parent table for all paragraphs in this batch
          for (let i = batchStart; i < batchEnd; i++) {
            const paragraph = allParagraphs.items[i];
            if (paragraph) {
              paragraph.load("text");
              paragraph.load("parentTableOrNullObject");
            }
          }
          await context.sync();
          
          // Load table IDs for paragraphs with parent tables
          for (let i = batchStart; i < batchEnd; i++) {
            const paragraph = allParagraphs.items[i];
            if (paragraph && paragraph.parentTableOrNullObject && !paragraph.parentTableOrNullObject.isNullObject) {
              paragraph.parentTableOrNullObject.load("id");
            }
          }
          await context.sync();
          
          // First filter to paragraphs that need processing
          const paragraphsToProcess = [];
          
          for (let i = batchStart; i < batchEnd; i++) {
            const paragraph = allParagraphs.items[i];
            if (!paragraph) continue;
            
            // Skip empty paragraphs
            if (!paragraph.text || paragraph.text.trim() === "") continue;
            
            // Skip if in a processed table
            let isInProcessedTable = false;
            if (paragraph.parentTableOrNullObject && !paragraph.parentTableOrNullObject.isNullObject) {
              if (processedTables.has(paragraph.parentTableOrNullObject.id)) {
                isInProcessedTable = true;
              }
            }
            
            if (isInProcessedTable) continue;
            
            // Check if this paragraph is already in a content control
            paragraph.load("parentContentControlOrNullObject");
            paragraphsToProcess.push(paragraph);
          }
          
          if (paragraphsToProcess.length === 0) {
            continue; // Skip to next batch if no paragraphs need processing
          }
          
          await context.sync();
          
          // Load tags on parent content controls
          for (let i = 0; i < paragraphsToProcess.length; i++) {
            const paragraph = paragraphsToProcess[i];
            if (paragraph && paragraph.parentContentControlOrNullObject && 
                !paragraph.parentContentControlOrNullObject.isNullObject) {
              paragraph.parentContentControlOrNullObject.load("tag");
            }
          }
          await context.sync();
          
          // Final processing - check if already in a para- content control
          for (let i = 0; i < paragraphsToProcess.length; i++) {
            const paragraph = paragraphsToProcess[i];
            
            // Check parent content control first
            let hasOurContentControl = false;
            let existingTag = "";
            
            if (paragraph.parentContentControlOrNullObject && 
                !paragraph.parentContentControlOrNullObject.isNullObject && 
                paragraph.parentContentControlOrNullObject.tag) {
              
              if (paragraph.parentContentControlOrNullObject.tag.startsWith("para-")) {
                hasOurContentControl = true;
                existingTag = paragraph.parentContentControlOrNullObject.tag;
                processedContentControls.add(existingTag);
                continue; // Already processed, skip to next paragraph
              }
            }
            
            // If not found yet, check paragraph's own content controls
            if (!hasOurContentControl) {
              try {
                // Load paragraph's content controls
                const paraContentControls = paragraph.contentControls;
                if (paraContentControls) {
                  paraContentControls.load("items");
                  await context.sync();
                
                  if (paraContentControls.items && paraContentControls.items.length > 0) {
                    // Load tags
                    for (let c = 0; c < paraContentControls.items.length; c++) {
                      if (paraContentControls.items[c]) {
                        paraContentControls.items[c].load("tag");
                      }
                    }
                    await context.sync();
                    
                    // Check for our tag format
                    for (let c = 0; c < paraContentControls.items.length; c++) {
                      const cc = paraContentControls.items[c];
                      if (cc && cc.tag && cc.tag.startsWith("para-")) {
                        hasOurContentControl = true;
                        existingTag = cc.tag;
                        processedContentControls.add(existingTag);
                        break;
                      }
                    }
                  }
                }
              } catch (e) {
                console.error("Error checking paragraph content controls:", e);
              }
            }
            
            // If still not found, insert a new content control
            if (!hasOurContentControl) {
              const uniqueId = `para-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
              try {
                const contentControl = paragraph.insertContentControl();
                if (contentControl) {
                  contentControl.tag = uniqueId;
                  contentControl.title = "Paragraph";
                  processedContentControls.add(uniqueId);
                }
              } catch (e) {
                console.error("Error inserting paragraph content control:", e);
              }
            }
          }
          
          // Sync at the end of processing this batch
          await context.sync();
        }
      }
      
      // PHASE 5: Gather all content controls for output
      console.log("Phase 5: Gathering content controls for output");
      
      // Get all content controls in the document
      const docContentControls = context.document.body.contentControls;
      docContentControls.load("items");
      await context.sync();
      
      if (!docContentControls.items || docContentControls.items.length === 0) {
        console.log("No content controls found in final gather phase");
        return { 
          elapsedTime: (Date.now() - startTime) / 1000,
          paragraphs: paragraphData
        };
      }
      
      // Create a temp array for all content controls before deduplication
      const tempContentControls = [];
      
      // Process content controls in batches
      const ccBatchSize = 20;
      for (let batchStart = 0; batchStart < docContentControls.items.length; batchStart += ccBatchSize) {
        const batchEnd = Math.min(batchStart + ccBatchSize, docContentControls.items.length);
        
        // Load tags and text for this batch
        for (let i = batchStart; i < batchEnd; i++) {
          const cc = docContentControls.items[i];
          if (cc) {
            cc.load("tag,text");
          }
        }
        await context.sync();
        
        // Filter to our content controls and add to output
        for (let i = batchStart; i < batchEnd; i++) {
          const cc = docContentControls.items[i];
          if (!cc || !cc.tag) continue;
          
          // Only process our tagged content controls
          if (cc.tag.startsWith('para-') || cc.tag.startsWith('table-')) {
            const isTable = cc.tag.startsWith('table-');
            
            // For tables, format a preview
            let displayText = cc.text || "";
            if (isTable) {
              displayText = "Table: " + displayText.substring(0, 50) + (displayText.length > 50 ? "..." : "");
            }
            
            // Add to our output array, will deduplicate later
            tempContentControls.push({
              id: cc.tag,
              text: displayText,
              isTable: isTable
            });
          }
        }
      }
      
      // Deduplicate based on ID
      const uniqueIds = new Set();
      for (const item of tempContentControls) {
        if (!uniqueIds.has(item.id)) {
          uniqueIds.add(item.id);
          paragraphData.push(item);
        }
      }
      
      console.log(`Document analysis complete. Found ${paragraphData.length} unique content controls.`);
      const elapsedTime = (Date.now() - startTime) / 1000; // in seconds
      
      return {
        elapsedTime,
        paragraphs: paragraphData
      };
    });
  } catch (error) {
    console.error("Error in processParagraphs:", error);
    return {
      elapsedTime: (Date.now() - startTime) / 1000,
      paragraphs: [],
      error: error.message
    };
  }
}

// Function to update paragraph content by ID
export async function updateParagraphById(paraId, newContent) {
  const startTime = Date.now();
  
  try {
    return await Word.run(async (context) => {
      // Search for content control with matching tag
      const contentControls = context.document.contentControls;
      contentControls.load("items,tag");
      await context.sync();
      
      // Find the content control with matching tag
      let found = false;
      for (let i = 0; i < contentControls.items.length; i++) {
        if (contentControls.items[i].tag === paraId) {
          // Update the content
          contentControls.items[i].insertText(newContent, Word.InsertLocation.replace);
          found = true;
          break;
        }
      }
      
      await context.sync();
      const elapsedTime = (Date.now() - startTime) / 1000; // in seconds
      
      return {
        success: found,
        elapsedTime
      };
    });
  } catch (error) {
    console.error("Error: " + error);
    return {
      success: false,
      elapsedTime: (Date.now() - startTime) / 1000,
      error: error.message
    };
  }
}

// Function to add a comment to a paragraph by ID
export async function addCommentToParagraph(paraId, commentText) {
  const startTime = Date.now();
  
  try {
    return await Word.run(async (context) => {
      // Search for content control with matching tag
      const contentControls = context.document.contentControls;
      contentControls.load("items,tag");
      await context.sync();
      
      // Find the content control with matching tag
      let found = false;
      for (let i = 0; i < contentControls.items.length; i++) {
        if (contentControls.items[i].tag === paraId) {
          // First, select the content control
          const range = contentControls.items[i].getRange();
          
          // Add a comment to the range
          range.comments.add(commentText);
          found = true;
          break;
        }
      }
      
      await context.sync();
      const elapsedTime = (Date.now() - startTime) / 1000; // in seconds
      
      return {
        success: found,
        elapsedTime
      };
    });
  } catch (error) {
    console.error("Error: " + error);
    return {
      success: false,
      elapsedTime: (Date.now() - startTime) / 1000,
      error: error.message
    };
  }
}

// Function to highlight/hover over a paragraph by ID
export async function highlightParagraph(paraId) {
  const startTime = Date.now();
  
  try {
    return await Word.run(async (context) => {
      // Search for content control with matching tag
      const contentControls = context.document.contentControls;
      contentControls.load("items,tag");
      await context.sync();
      
      // Find the content control with matching tag
      let found = false;
      for (let i = 0; i < contentControls.items.length; i++) {
        if (contentControls.items[i].tag === paraId) {
          // Highlight the paragraph (change background color to yellow)
          contentControls.items[i].font.highlightColor = 'yellow';
          contentControls.items[i].select();
          found = true;
          break;
        }
      }
      
      await context.sync();
      const elapsedTime = (Date.now() - startTime) / 1000; // in seconds
      
      return {
        success: found,
        elapsedTime
      };
    });
  } catch (error) {
    console.error("Error: " + error);
    return {
      success: false,
      elapsedTime: (Date.now() - startTime) / 1000,
      error: error.message
    };
  }
}

// Function to scroll to a paragraph by ID without highlighting it
export async function scrollToParagraph(paraId) {
  const startTime = Date.now();
  
  try {
    return await Word.run(async (context) => {
      // Search for content control with matching tag
      const contentControls = context.document.contentControls;
      contentControls.load("items,tag");
      await context.sync();
      
      // Find the content control with matching tag
      let found = false;
      for (let i = 0; i < contentControls.items.length; i++) {
        if (contentControls.items[i].tag === paraId) {
          // Get the range of the content control
          const range = contentControls.items[i].getRange();
          
          // Scroll into view without highlighting
          range.select();
          range.scrollIntoView();
          found = true;
          break;
        }
      }
      
      await context.sync();
      const elapsedTime = (Date.now() - startTime) / 1000; // in seconds
      
      return {
        success: found,
        elapsedTime
      };
    });
  } catch (error) {
    console.error("Error: " + error);
    return {
      success: false,
      elapsedTime: (Date.now() - startTime) / 1000,
      error: error.message
    };
  }
}

// Original insert text function for reference
export async function insertText(text) {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
