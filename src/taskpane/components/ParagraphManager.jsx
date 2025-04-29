import * as React from "react";
import { useState, useEffect } from "react";
import { 
  Button, Field, Textarea, Input, makeStyles, 
  Label, tokens, Spinner, TabList, Tab, 
  Card, CardHeader, Text,
  mergeClasses, Badge, Divider,
  Tooltip
} from "@fluentui/react-components";
import { 
  Document24Regular, 
  DocumentEdit24Regular, 
  Comment24Regular, 
  TextBulletListSquare24Regular, 
  ArrowClockwise24Regular,
  DocumentSearch24Regular,
  ArrowRight24Regular,
  Table24Regular
} from "@fluentui/react-icons";
import PropTypes from "prop-types";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  analysisButton: {
    maxWidth: "200px",
    marginBottom: "12px",
    transition: "all 0.2s ease",
  },
  analyzeButtonContent: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  timingInfo: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
    display: "flex",
    alignItems: "center",
    gap: "4px",
    marginBottom: "4px",
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "4px",
    marginTop: "0",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  paragraphListContainer: {
    height: "350px",
    display: "flex",
    flexDirection: "column",
  },
  paragraphList: {
    height: "100%",
    overflowY: "auto",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "4px",
    padding: "4px",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  paragraph: {
    padding: "4px 6px",
    borderRadius: "2px",
    marginBottom: "2px",
    transition: "all 0.2s ease",
    cursor: "pointer",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground2,
    },
    "&:last-child": {
      borderBottom: "none",
    }
  },
  paragraphId: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase100,
    display: "flex",
    alignItems: "center",
    gap: "4px",
  },
  paragraphText: {
    fontSize: tokens.fontSizeBase200,
    marginTop: "2px",
    overflow: "hidden",
    textOverflow: "ellipsis",
    display: "-webkit-box",
    WebkitLineClamp: 2,
    WebkitBoxOrient: "vertical",
  },
  paragraphActions: {
    display: "flex",
    justifyContent: "flex-end",
    gap: "8px",
    marginTop: "2px",
  },
  actionIcon: {
    cursor: "pointer",
    color: tokens.colorBrandBackground,
    fontSize: tokens.fontSizeBase200,
    "&:hover": {
      color: tokens.colorBrandBackgroundHover,
    },
  },
  operationsSection: {
    backgroundColor: tokens.colorNeutralBackground1,
    padding: "8px 10px",
    borderRadius: "4px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  tabList: {
    marginBottom: "8px",
  },
  operationPanel: {
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: "4px",
  },
  actionButton: {
    minWidth: "120px",
  },
  operationCard: {
    marginTop: "12px",
    boxShadow: "0 2px 4px rgba(0,0,0,0.05)",
  },
  inputField: {
    marginBottom: "8px",
  },
  cardSuccess: {
    borderLeft: `4px solid ${tokens.colorStatusSuccessBackground2}`,
  },
  cardError: {
    borderLeft: `4px solid ${tokens.colorStatusDangerBackground2}`,
  },
  spinner: {
    padding: "12px",
    display: "flex",
    justifyContent: "center",
  },
  badge: {
    marginLeft: "8px",
  },
  divider: {
    margin: "8px 0",
  },
  resultsContainer: {
    display: "flex",
    flexDirection: "column",
  }
});

const ParagraphManager = (props) => {
  const [analyzing, setAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState(null);
  const [selectedTab, setSelectedTab] = useState("replace");
  const [paraId, setParaId] = useState("");
  const [newContent, setNewContent] = useState("");
  const [commentText, setCommentText] = useState("");
  const [operationResult, setOperationResult] = useState(null);
  
  const styles = useStyles();

  // Clear operation result after 5 seconds
  useEffect(() => {
    if (operationResult && !operationResult.inProgress) {
      const timer = setTimeout(() => {
        setOperationResult(null);
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [operationResult]);

  const handleAnalyzeDocument = async () => {
    setAnalyzing(true);
    setAnalysisResult(null);
    setOperationResult(null);
    
    try {
      const result = await props.processParagraphs();
      setAnalysisResult(result);
    } catch (error) {
      console.error("Analysis error:", error);
    } finally {
      setAnalyzing(false);
    }
  };

  const handleTabSelect = (event, data) => {
    setSelectedTab(data.value);
    setOperationResult(null);
  };

  const handleReplaceContent = async () => {
    if (!paraId || !newContent) return;
    
    setOperationResult({ inProgress: true });
    const result = await props.updateParagraphById(paraId, newContent);
    setOperationResult({ 
      type: 'replace',
      ...result 
    });
  };

  const handleAddComment = async () => {
    if (!paraId || !commentText) return;
    
    setOperationResult({ inProgress: true });
    const result = await props.addCommentToParagraph(paraId, commentText);
    setOperationResult({ 
      type: 'comment',
      ...result 
    });
  };

  const handleHighlight = async () => {
    if (!paraId) return;
    
    setOperationResult({ inProgress: true });
    const result = await props.highlightParagraph(paraId);
    setOperationResult({ 
      type: 'highlight',
      ...result 
    });
  };

  // New function to scroll to and select paragraph without highlighting
  const scrollAndSelectParagraph = async (id) => {
    setParaId(id);
    
    try {
      // Create a function in props that exposes the scrollIntoView functionality
      // This is expected to be implemented in the main taskpane.js file
      await props.scrollToParagraph(id);
      
      const tempResult = {
        success: true,
        type: 'navigate',
        elapsedTime: 0
      };
      
      setOperationResult({ 
        ...tempResult,
        inProgress: false
      });
    } catch (error) {
      console.error("Navigation error:", error);
      
      // Fallback to the highlight function if scrollToParagraph isn't available
      if (!props.scrollToParagraph) {
        try {
          await props.highlightParagraph(id);
        } catch (fallbackError) {
          console.error("Fallback navigation error:", fallbackError);
        }
      }
    }
  };

  return (
    <div className={styles.container}>
      {/* Analysis Button */}
      <Button 
        appearance="primary" 
        onClick={handleAnalyzeDocument} 
        disabled={analyzing}
        className={styles.analysisButton}
        icon={<DocumentSearch24Regular />}
      >
        {analyzing ? "Analyzing..." : "Analyze Document"}
      </Button>
      
      {analyzing && 
        <div className={styles.spinner}>
          <Spinner size="small" label="Analyzing document..." />
        </div>
      }
      
      {/* Analysis Results */}
      {analysisResult && (
        <div className={styles.resultsContainer}>
          <div className={styles.timingInfo}>
            <ArrowClockwise24Regular />
            Analysis complete in {analysisResult.elapsedTime.toFixed(2)} seconds
          </div>
          
          <Text className={styles.sectionTitle}>
            <span>
              Document Paragraphs
              <Badge 
                appearance="filled" 
                color="informative" 
                className={styles.badge}
              >
                {analysisResult.paragraphs.length}
              </Badge>
            </span>
          </Text>
          
          <div className={styles.paragraphListContainer}>
            <div className={styles.paragraphList}>
              {analysisResult.paragraphs.map((para, index) => (
                <div 
                  key={para.id} 
                  className={styles.paragraph}
                  onClick={() => scrollAndSelectParagraph(para.id)}
                >
                  <div className={styles.paragraphId}>
                    {para.isTable ? (
                      <Table24Regular style={{fontSize: '14px'}} />
                    ) : (
                      <Document24Regular style={{fontSize: '14px'}} />
                    )}
                    ID: {para.id}
                  </div>
                  <div className={styles.paragraphText}>{para.text}</div>
                  <div className={styles.paragraphActions}>
                    <Tooltip content="Go to paragraph in document" relationship="label">
                      <ArrowRight24Regular 
                        className={styles.actionIcon} 
                        style={{fontSize: '14px'}}
                        onClick={(e) => {
                          e.stopPropagation();
                          scrollAndSelectParagraph(para.id);
                        }} 
                      />
                    </Tooltip>
                  </div>
                </div>
              ))}
            </div>
          </div>
          
          <Divider className={styles.divider} />
          
          {/* Operations Section */}
          <Text className={styles.sectionTitle}>Paragraph Operations</Text>
          <div className={styles.operationsSection}>
            <Field className={styles.inputField} label="Paragraph ID">
              <Input 
                value={paraId} 
                onChange={(e) => setParaId(e.target.value)} 
                placeholder="Enter paragraph ID or click a paragraph above"
              />
            </Field>
            
            <TabList 
              selectedValue={selectedTab} 
              onTabSelect={handleTabSelect}
              className={styles.tabList}
              appearance="subtle"
            >
              <Tab value="replace" icon={<DocumentEdit24Regular />}>Replace</Tab>
              <Tab value="comment" icon={<Comment24Regular />}>Comment</Tab>
              <Tab value="highlight" icon={<TextBulletListSquare24Regular />}>Highlight</Tab>
            </TabList>
            
            <div className={styles.operationPanel}>
              {selectedTab === "replace" && (
                <div>
                  <Field className={styles.inputField} label="New Content">
                    <Textarea
                      value={newContent}
                      onChange={(e) => setNewContent(e.target.value)}
                      placeholder="Enter new content for the paragraph"
                      resize="vertical"
                      style={{minHeight: "60px", maxHeight: "100px"}}
                    />
                  </Field>
                  <Button 
                    appearance="primary" 
                    onClick={handleReplaceContent}
                    disabled={!paraId || !newContent || (operationResult && operationResult.inProgress)}
                    className={styles.actionButton}
                    icon={<DocumentEdit24Regular />}
                  >
                    Replace Content
                  </Button>
                </div>
              )}
              
              {selectedTab === "comment" && (
                <div>
                  <Field className={styles.inputField} label="Comment">
                    <Textarea
                      value={commentText}
                      onChange={(e) => setCommentText(e.target.value)}
                      placeholder="Enter comment text"
                      resize="vertical"
                      style={{minHeight: "60px", maxHeight: "100px"}}
                    />
                  </Field>
                  <Button 
                    appearance="primary" 
                    onClick={handleAddComment}
                    disabled={!paraId || !commentText || (operationResult && operationResult.inProgress)}
                    className={styles.actionButton}
                    icon={<Comment24Regular />}
                  >
                    Add Comment
                  </Button>
                </div>
              )}
              
              {selectedTab === "highlight" && (
                <div>
                  <Button 
                    appearance="primary" 
                    onClick={handleHighlight}
                    disabled={!paraId || (operationResult && operationResult.inProgress)}
                    className={styles.actionButton}
                    icon={<TextBulletListSquare24Regular />}
                  >
                    Highlight Paragraph
                  </Button>
                </div>
              )}
            </div>
          </div>
          
          {/* Operation Result */}
          {operationResult && operationResult.inProgress && (
            <div className={styles.spinner}>
              <Spinner size="tiny" label="Processing..." labelPosition="after" />
            </div>
          )}
          
          {operationResult && !operationResult.inProgress && (
            <Card 
              className={mergeClasses(
                styles.operationCard, 
                operationResult.success ? styles.cardSuccess : styles.cardError
              )}
            >
              <CardHeader header={
                <Text weight="semibold">
                  {operationResult.type === 'replace' ? 'Replace' : 
                   operationResult.type === 'comment' ? 'Comment' : 
                   operationResult.type === 'navigate' ? 'Navigation' : 'Highlight'} Result
                </Text>
              } />
              <div style={{ padding: '0 16px 16px' }}>
                <Text>
                  {operationResult.type === 'navigate' 
                    ? 'Successfully navigated to paragraph in document.'
                    : operationResult.success 
                      ? `Success! Operation completed in ${operationResult.elapsedTime.toFixed(2)} seconds.` 
                      : `Operation failed. ${operationResult.error || 'Paragraph ID not found.'}`}
                </Text>
              </div>
            </Card>
          )}
        </div>
      )}
    </div>
  );
};

ParagraphManager.propTypes = {
  processParagraphs: PropTypes.func.isRequired,
  updateParagraphById: PropTypes.func.isRequired,
  addCommentToParagraph: PropTypes.func.isRequired,
  highlightParagraph: PropTypes.func.isRequired,
  scrollToParagraph: PropTypes.func,
};

export default ParagraphManager; 