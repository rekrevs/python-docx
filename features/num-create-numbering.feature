Feature: Create numbering definitions
  In order to programmatically create numbered and bulleted lists
  As a developer using python-docx
  I need to create abstract numbering definitions and apply them to paragraphs


  Scenario: Create a new numbering definition and apply it to paragraphs
    Given a new document
     When I create a decimal numbering definition
      And I apply the numbering to a paragraph
     Then the paragraph has numbering applied

  Scenario: Created numbering survives save and reload
    Given a new document with a custom numbering definition
     When I save and reload the document
     Then the custom numbering definition is preserved
