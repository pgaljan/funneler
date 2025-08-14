name: 📋 Code of Conduct Violation Report
description: Report behavior that violates our community standards
title: "[CODE OF CONDUCT] Brief description of the issue"
labels: ["code-of-conduct", "needs-review"]
assignees:
  - maintainer-username
body:
  - type: markdown
    attributes:
      value: |
        ## Important Notice
        
        Thank you for helping us maintain a safe and welcoming community. All reports are taken seriously and will be handled confidentially.
        
        **For urgent safety concerns or harassment, please email us directly at: conduct@projectname.org**
  
  - type: textarea
    id: incident-description
    attributes:
      label: What happened?
      description: Please provide a clear description of the incident
      placeholder: Describe what occurred, when it happened, and who was involved...
    validations:
      required: true
  
  - type: input
    id: location
    attributes:
      label: Where did this occur?
      description: Location of the incident (Issue #, PR #, Discussion, etc.)
      placeholder: "Example: Issue #123, PR #456, or Discord channel"
    validations:
      required: true
  
  - type: checkboxes
    id: violation-type
    attributes:
      label: Type of violation (check all that apply)
      options:
        - label: Harassment or bullying
        - label: Discriminatory language or behavior
        - label: Spam or off-topic content
        - label: Personal attacks or insults
        - label: Sharing private information
        - label: Other (please describe above)
  
  - type: dropdown
    id: urgency
    attributes:
      label: Urgency Level
      description: How urgent is this matter?
      options:
        - Low - Can be addressed in normal course
        - Medium - Should be addressed within a few days
        - High - Needs immediate attention
        - Critical - Safety concern requiring urgent action
    validations:
      required: true
  
  - type: checkboxes
    id: reporter-info
    attributes:
      label: Reporter Information
      options:
        - label: I am directly affected by this incident
        - label: I witnessed this incident
        - label: I am reporting on behalf of someone else
        - label: I prefer to remain anonymous in any follow-up
  
  - type: textarea
    id: additional-info
    attributes:
      label: Additional Information
      description: Any other details that might be helpful
      placeholder: Links to relevant content, screenshots (with personal info removed), context, etc.
  
  - type: markdown
    attributes:
      value: |
        ## What happens next?
        
        1. A maintainer will review this report within 24-48 hours
        2. We may reach out for additional information if needed
        3. We will investigate the matter confidentially
        4. Appropriate action will be taken based on our Code of Conduct
        5. You will be notified of the resolution (if requested)
        
        **This issue will be made private automatically to protect all parties involved.**