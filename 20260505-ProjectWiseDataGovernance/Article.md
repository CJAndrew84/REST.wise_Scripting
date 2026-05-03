## Why ProjectWise Templates Don’t Scale — And What to Do About It

At Bentley Illuminate 2026, one of the sessions by Benito Perez Galan focused on automating ProjectWise project creation.

The approach was straightforward:
- Define your project in Excel  
- Use a script to build the ProjectWise structure  
- Reduce manual setup effort  

It’s a solid step forward—and one many of us have been moving toward for years.

There were also some practical elements in the session I’ll take away and use. In particular, the way Benito structured disciplines, their associated folders, and the user lists tied to them is something I’ll be adopting going forward. It’s a cleaner way of thinking about how discipline-based access and structure should be set up.

I’ve been using a variation of this approach since around 2018, when I was working as a Lead ProjectWise Admin for the Highways business in the UK at AECOM (known internally as a GA role).

Driving template setup is a clear improvement over manual creation in ProjectWise Administrator.

What Benito’s session showed was project creation from Excel acting as the template. In reality, most of us already have ways to copy an existing template and deploy it as a project.

Where his approach adds value is in reducing the overhead—particularly around discipline folder structures and associated user lists—by defining that directly in the template.

That’s something I’ve worked with before. The foundations were already in place from predecessors, and colleagues built a UI to allow project requests to specify the disciplines required. The deployment script then only creates what’s needed for that project, rather than everything by default.

That works well for setup.

But the part I’m more interested in now is different.

> It’s less about how we create templates, and more about how we document, govern, and track them over time.

The aim is simple:

> If someone joins the business tomorrow, they should be able to understand what the current template is, how it got there, and why it’s been set up that way.

That’s the gap I think still exists.

And it’s what I’ve been focusing on more recently.

---

[IMAGE PLACEHOLDER: Benito’s Illuminate slide showing Excel → ProjectWise automation]

---

## The Real Issue Isn’t Setup… It’s Control

Most organisations still treat ProjectWise templates as:
- Static  
- Manual to manage  
- Difficult to compare  
- Hard to govern at scale  

That leads to familiar issues:
- Inconsistent project setups  
- Template drift between environments (DEV / UAT / PROD)  
- Limited visibility of what’s actually been deployed  
- Governance that is reactive rather than proactive  

So while automating setup is useful, it doesn’t answer a more important question:

> How do you control, validate, and continuously improve your ProjectWise environment over time?

---

Before going further, it’s worth asking a simple question:

> How many of you reading this have inherited a ProjectWise setup you didn’t build?

Something that:
- You don’t fully understand  
- May have some documentation  
- Probably doesn’t  

That’s a fairly common situation.

And it’s where this approach starts to help.

---

## A Different Way to Look at It

The shift I’ve been working toward is this:

> Stop treating templates as configuration, and start treating them as structured data.

Once you do that, a few things become possible:
- You can extract your current ProjectWise setup into a structured format  
- You can compare environments properly  
- You can promote changes through DEV → UAT → PROD  
- You can validate compliance against governance rules  

You move from:
- “Did we set this project up correctly?”  

to:
- “Is our environment consistently aligned with how we say it should be configured?”

---

## My Approach: Treat the Setup as Data

Instead of starting with a template and trying to make it fit every project, I extract the actual configuration out of ProjectWise into a structured format.

At the moment, that’s Excel.

That includes:
- Folder structures  
- Workflows  
- Security (User Lists / Groups)  
- Environments and attributes  
- Document codes and metadata  
- Views, rules, lookup tables and more  

The key point is:

> This represents what is actually deployed—not what we think is deployed.

---

## From Workbook to Data Pipeline

The first version of this approach exports everything into a single Excel workbook.

That gives one structured view of the setup across:
- Environments and attributes  
- Workflows and states  
- Views  
- Attribute exchange rules  
- Rich project definitions  
- Lookup tables  
- WRE rules  
- Folder structures  
- Access control  
- Project resources  
- Disciplines and user lists  

That alone is useful—it makes the setup visible.

But Excel isn’t the end goal.

The next step is exporting the same information to JSON.

That matters because JSON can be:
- Compared  
- Versioned  
- Queried  
- Used in pipelines  
- Fed into validation and reporting  

> Excel makes the setup understandable.  
> JSON makes it automatable.

---

[IMAGE PLACEHOLDER: Excel workbook with multiple configuration sheets]

[IMAGE PLACEHOLDER: JSON diff or version comparison in repository]

---

## How This Works in Practice

The process I’m working toward looks like this:

### 1. Define the Template in Excel
The Excel workbook becomes the structured definition of the setup.

---

### 2. Deploy to Development
PowerShell scripts deploy that definition into a Development datasource.

This is where the template is built and tested properly.

---

### 3. Capture the Deployed State (JSON)
Once it’s working, the deployed configuration is exported to JSON and stored in a repository (Azure DevOps in my case, GitHub would work the same).

> This captures what was actually deployed—not just what was intended.

---

### 4. Promote to UAT
The template is deployed into a UAT datasource for user validation.

Workflows, metadata, and usability get tested in a real context.

---

### 5. Feed Changes Back Properly
Any changes from UAT or Production are not patched directly.

Instead:
- Update the Excel definition  
- Redeploy to Development  
- Re-test  
- Re-export JSON  
- Commit to the repository  

---

### 6. Deploy to Production
Once validated, the same process pushes the template to Production.

---

### 7. Track Change Over Time
Because each JSON export is version-controlled:
- You can see exactly what changed  
- You can track when and why it changed  
- You can compare environments properly  

Git becomes useful here—not for code, but for configuration history.

---

[IMAGE PLACEHOLDER: DEV → UAT → PROD flow diagram with Excel + JSON + Git]

---

## Where This Is Heading

Once the setup is structured and version-controlled, you can start adding control around how it evolves.

### Pipelines
Formalise the process:
- Define → Deploy → Capture → Promote  

Make it repeatable and traceable.

---

### Dashboards
Surface the state of the environment:
- Which templates align with governance  
- Where environments differ  
- What has changed over time  

---

### Automated Validation
Start checking the setup against expected rules:
- Naming conventions  
- Required attributes  
- Workflow structures  
- Folder hierarchies  
- Access control  

When something falls outside of that, it gets flagged.

---

## Handling Exceptions Properly

Not everything will fit governance—and that’s fine.

The aim isn’t to block work.

It’s to make deviations visible.

If something is built outside the expected structure:

> The administrator should be prompted to document the reason for the exemption.

That way:
- Exceptions are recorded  
- Decisions are visible  
- Patterns can be reviewed over time  

---

## What This Changes

This isn’t about making setup faster.

It’s about changing how ProjectWise is managed:

- From templates → to data-driven definitions  
- From manual checks → to repeatable validation  
- From drift → to controlled alignment  

And importantly:

> From reacting to issues  
> to being able to see them before they become problems.

---

## Reality Check

This isn’t a finished solution.

It still depends on:
- Clean, structured data  
- Discipline in how changes are made  
- A clear definition of what “good” looks like  

Excel is just a starting point—not the end state.

---

## Final Thought

Benito’s session showed how we can automate project creation.

This is really about the next step:

> How do we control, validate, and continuously improve what we’ve created?

One practical example of where this can go:

I recently took one of the exported Excel templates and asked AI to review it.

It picked up something that isn’t always obvious when you’re working inside ProjectWise—cascading triggers within environments.

On paper, they looked fine.

In practice, the AI flagged that advancing even a relatively small number of files through a workflow could cause performance bottlenecks because of how those triggers interact.

That’s the kind of issue that’s easy to miss, especially as templates grow over time.

It’s not about replacing experience—it’s about having another way to interrogate the setup once it’s structured and visible.

---

[IMAGE PLACEHOLDER: Screenshot of AI response highlighting cascading triggers / performance concern]

---

And this is where it becomes more useful day-to-day.

Once the setup exists in a structured format—whether that’s Excel or JSON—you can start to use AI to:

- Explain how the template is structured  
- Describe workflows and state transitions  
- Highlight dependencies and triggers  
- Generate simple flow diagrams showing how things behave  

You’re no longer limited to digging through ProjectWise Administrator trying to piece things together.

And more importantly:

> You’re no longer in the position of saying  
> “I don’t support that—I didn’t build it.”

---

## Interested to hear how others are approaching this

Are you:
- Still relying on static templates?  
- Using structured inputs for setup?  
- Doing anything around comparison or validation across environments?  

Would be good to compare approaches—especially where things start to scale.