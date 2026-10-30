import streamlit as st
import re
import io
import random
from pypdf import PdfReader
from docx import Document
from pptx import Presentation

# --------------------------------------------------
# PAGE CONFIGURATION
# --------------------------------------------------

st.set_page_config(
    page_title="InterviewPreps | Interview Practice",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --------------------------------------------------
# CUSTOM CSS - PREMIUM UI
# --------------------------------------------------

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');

:root {
    --primary: #5146e5;
    --primary-dark: #4035c7;
    --text: #172033;
    --muted: #64748b;
    --background: #f6f7fb;
    --card: #ffffff;
    --border: #e5e7eb;
}

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.stApp {
    background: var(--background);
    color: var(--text);
}

[data-testid="stAppViewContainer"] {
    background: var(--background);
}

[data-testid="stHeader"] {
    background: transparent;
}

.block-container {
    max-width: 1150px;
    padding-top: 1.5rem;
    padding-bottom: 3rem;
}

/* Normal text */
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stWidgetLabel"] p,
.stCaption,
label {
    color: var(--text);
}

/* Headings */
h1, h2, h3, h4 {
    font-family: 'Space Grotesk', sans-serif !important;
    color: var(--text) !important;
    overflow: visible !important;
    overflow-wrap: anywhere !important;
    line-height: 1.3 !important;
    height: auto !important;
}

/* Brand header */
.brand-header {
    display: flex;
    align-items: center;
    gap: 12px;
    width: 100%;
    min-width: 0;
    padding: 6px 0 20px 0;
    overflow: visible;
}

.brand-logo {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 48px;
    min-width: 48px;
    height: 48px;
    border-radius: 14px;
    background: var(--primary);
    color: #ffffff !important;
    font-size: 25px;
}

.brand-name {
    display: flex;
    flex-wrap: wrap;
    align-items: baseline;
    font-family: 'Space Grotesk', sans-serif;
    font-size: clamp(23px, 3.4vw, 34px);
    font-weight: 700;
    line-height: 1.35;
    overflow: visible;
}

.brand-first {
    color: #172033 !important;
}

.brand-last {
    color: var(--primary) !important;
}

/* Hero */
.hero {
    padding: 42px 34px;
    border-radius: 24px;
    background: linear-gradient(135deg, #5146e5 0%, #7569f5 65%, #8a80ff 100%);
    margin-bottom: 25px;
    overflow: visible;
}

.hero h1 {
    color: #ffffff !important;
    font-size: clamp(30px, 4.5vw, 48px) !important;
    line-height: 1.2 !important;
    margin-bottom: 12px;
}

.hero p {
    color: #eeedff !important;
    font-size: 16px;
    line-height: 1.8;
    max-width: 740px;
}

.hero-tag {
    display: inline-block;
    background: rgba(255,255,255,0.16);
    color: #ffffff !important;
    border: 1px solid rgba(255,255,255,0.25);
    border-radius: 30px;
    padding: 6px 12px;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.4px;
    margin-bottom: 15px;
}

/* Cards */
.info-card {
    background: #ffffff;
    border: 1px solid #e6e8f0;
    border-radius: 18px;
    padding: 23px;
    margin-bottom: 15px;
    box-shadow: 0 4px 15px rgba(23, 32, 51, 0.035);
}

.info-card h3 {
    margin-top: 0;
}

.info-card p {
    color: #64748b;
    line-height: 1.7;
}

.step-label {
    color: var(--primary) !important;
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
}

/* Buttons */
.stButton > button,
.stDownloadButton > button {
    background-color: #ffffff !important;
    color: #5146e5 !important;
    border: 1px solid #5146e5 !important;
    border-radius: 11px !important;
    min-height: 44px;
    font-weight: 700 !important;
    opacity: 1 !important;
    transition: all 0.2s ease;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background-color: #5146e5 !important;
    color: #ffffff !important;
    border-color: #5146e5 !important;
}

.stButton > button:focus,
.stDownloadButton > button:focus {
    box-shadow: 0 0 0 3px rgba(81, 70, 229, 0.18) !important;
}

/* Primary buttons */
.stButton > button[kind="primary"],
.stButton > button[kind="primaryFormSubmit"] {
    background-color: #5146e5 !important;
    color: #ffffff !important;
    border: 1px solid #5146e5 !important;
}

.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primaryFormSubmit"]:hover {
    background-color: #4035c7 !important;
    color: #ffffff !important;
}

/* Text inputs */
.stTextInput input,
.stTextArea textarea,
[data-baseweb="select"] > div {
    color: #172033 !important;
    background-color: #ffffff !important;
    border-color: #d9deea !important;
    border-radius: 10px !important;
}

.stTextInput input::placeholder,
.stTextArea textarea::placeholder {
    color: #7b8498 !important;
    opacity: 1 !important;
}

[data-baseweb="select"] * {
    color: #172033 !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: #ffffff;
    border: 1px dashed #b9b5f5;
    border-radius: 14px;
    padding: 12px;
}

[data-testid="stFileUploader"] section {
    background: #ffffff;
}

[data-testid="stFileUploader"] button {
    color: #5146e5 !important;
    border-color: #5146e5 !important;
}

/* Progress */
.stProgress > div > div > div > div {
    background-color: #5146e5;
}

/* Alerts */
[data-testid="stAlert"] {
    border-radius: 12px;
}

/* Divider */
hr {
    border-color: #e6e8f0;
}

/* Footer */
.footer {
    text-align: center;
    color: #7b8498 !important;
    font-size: 13px;
    padding-top: 28px;
}

/* Small screens */
@media (max-width: 640px) {
    .block-container {
        padding-left: 1rem;
        padding-right: 1rem;
    }

    .hero {
        padding: 28px 22px;
        border-radius: 18px;
    }

    .brand-header {
        gap: 9px;
    }

    .brand-logo {
        width: 42px;
        min-width: 42px;
        height: 42px;
    }
}
</style>
""", unsafe_allow_html=True)


# --------------------------------------------------
# INTERVIEW QUESTION BANK
# --------------------------------------------------

QUESTION_BANK = {
    "Python Developer": [
        {
            "q": "What are the main features of Python?",
            "a": "Python is a high-level, interpreted, dynamically typed programming language. Its main features include readable syntax, extensive libraries, cross-platform support, object-oriented programming, and automatic memory management.",
            "keywords": ["high-level", "interpreted", "libraries", "readable", "object-oriented"]
        },
        {
            "q": "What is the difference between a list and a tuple?",
            "a": "A list is mutable, which means its elements can be changed after creation. A tuple is immutable. Lists use square brackets, while tuples usually use parentheses. Lists are suitable for changing collections, while tuples are useful for fixed data.",
            "keywords": ["mutable", "immutable", "list", "tuple", "brackets"]
        },
        {
            "q": "What are Python dictionaries?",
            "a": "A dictionary stores data as key-value pairs. Keys must be unique and hashable. Dictionaries are mutable and are useful for fast lookups by key.",
            "keywords": ["key-value", "unique", "mutable", "lookup"]
        },
        {
            "q": "What is exception handling in Python?",
            "a": "Exception handling manages runtime errors using try, except, else, and finally. The try block contains risky code, except handles errors, else runs if no exception occurs, and finally runs whether an exception occurs or not.",
            "keywords": ["try", "except", "finally", "errors"]
        },
        {
            "q": "Explain object-oriented programming in Python.",
            "a": "Object-oriented programming organizes code around classes and objects. Its main principles are encapsulation, inheritance, polymorphism, and abstraction. It helps make software reusable, modular, and easier to maintain.",
            "keywords": ["classes", "objects", "inheritance", "polymorphism", "encapsulation"]
        },
        {
            "q": "What is the difference between == and is in Python?",
            "a": "The == operator compares whether two objects have equal values. The is operator checks whether both references point to the same object in memory.",
            "keywords": ["value", "identity", "object", "memory"]
        },
        {
            "q": "What are Python functions and why are they useful?",
            "a": "A function is a reusable block of code defined using def. Functions accept parameters, can return results, reduce repetition, and make programs easier to test and maintain.",
            "keywords": ["def", "reusable", "parameters", "return", "maintain"]
        },
        {
            "q": "How would you debug a Python program?",
            "a": "I would reproduce the issue, read the error message and traceback, inspect relevant variable values, and isolate the faulty section. I would use print statements or a debugger, fix the root cause, and test normal and edge cases.",
            "keywords": ["error", "traceback", "debugger", "root cause", "test"]
        }
    ],

    "Data Analyst": [
        {
            "q": "What does a data analyst do?",
            "a": "A data analyst collects, cleans, transforms, and analyzes data to identify patterns and support decisions. Common tools include SQL, Excel, Python, and visualization tools such as Power BI or Tableau.",
            "keywords": ["collects", "cleans", "analyzes", "SQL", "visualization"]
        },
        {
            "q": "What is data cleaning?",
            "a": "Data cleaning improves data quality by handling missing values, duplicates, inconsistent formats, incorrect entries, and outliers. I would inspect the dataset first and choose an appropriate treatment based on the data and business context.",
            "keywords": ["missing", "duplicates", "inconsistent", "outliers", "quality"]
        },
        {
            "q": "What is the difference between mean, median, and mode?",
            "a": "Mean is the sum of values divided by the number of values. Median is the middle value after sorting. Mode is the most frequently occurring value. Median is often more robust when data contains extreme outliers.",
            "keywords": ["average", "middle", "frequent", "outliers"]
        },
        {
            "q": "What is exploratory data analysis?",
            "a": "Exploratory Data Analysis, or EDA, is the process of understanding a dataset using summary statistics and visualizations. It helps identify distributions, missing values, outliers, relationships, and patterns before modeling.",
            "keywords": ["summary statistics", "visualizations", "missing", "outliers", "patterns"]
        },
        {
            "q": "What is the difference between INNER JOIN and LEFT JOIN?",
            "a": "INNER JOIN returns rows with matching values in both tables. LEFT JOIN returns every row from the left table and matching rows from the right table; unmatched right-side values appear as NULL.",
            "keywords": ["matching", "left table", "NULL", "INNER JOIN", "LEFT JOIN"]
        },
        {
            "q": "How would you explain a dashboard to a non-technical stakeholder?",
            "a": "I would begin with the business question, highlight the most important metrics, use clear charts and labels, and explain the main trends in plain language. I would focus on actionable findings rather than technical terminology.",
            "keywords": ["business", "metrics", "trends", "clear", "actionable"]
        },
        {
            "q": "What is correlation, and does it mean causation?",
            "a": "Correlation measures how two variables move together. It does not prove that one variable causes the other because confounding variables or coincidence may explain the relationship. Further analysis or experiments may be required.",
            "keywords": ["relationship", "variables", "causation", "confounding"]
        },
        {
            "q": "How do you handle missing values in a dataset?",
            "a": "I first measure how many values are missing and examine why they are missing. Depending on the context, I may remove affected rows, impute values using suitable statistics, or use a method appropriate to the data type.",
            "keywords": ["inspect", "missing", "impute", "remove", "context"]
        }
    ],

    "Cloud Data Engineer": [
        {
            "q": "What does a cloud data engineer do?",
            "a": "A cloud data engineer builds and maintains systems that collect, transform, store, and deliver data using cloud infrastructure. Responsibilities include data pipelines, data quality, storage, orchestration, monitoring, security, and cost optimization.",
            "keywords": ["pipelines", "transform", "cloud", "storage", "monitoring"]
        },
        {
            "q": "What is ETL?",
            "a": "ETL stands for Extract, Transform, Load. Data is extracted from sources, transformed to meet quality and business requirements, and loaded into a target system such as a data warehouse.",
            "keywords": ["extract", "transform", "load", "source", "warehouse"]
        },
        {
            "q": "What is the difference between a data lake and a data warehouse?",
            "a": "A data lake can store large amounts of raw data in different formats. A data warehouse generally stores structured, curated data optimized for analytics and reporting. Both can be used together in a data platform.",
            "keywords": ["raw data", "structured", "analytics", "storage", "formats"]
        },
        {
            "q": "What is a data pipeline?",
            "a": "A data pipeline is a sequence of steps that moves and processes data from sources to destinations. It may include ingestion, validation, transformation, storage, scheduling, monitoring, and error handling.",
            "keywords": ["ingestion", "transformation", "storage", "monitoring", "validation"]
        },
        {
            "q": "How would you handle a failed data pipeline?",
            "a": "I would check logs and alerts, identify the failed stage, inspect the input data and dependencies, and determine the root cause. After fixing it, I would rerun the affected task safely, verify the output, and add monitoring to prevent recurrence.",
            "keywords": ["logs", "root cause", "dependencies", "rerun", "monitoring"]
        },
        {
            "q": "What is SQL used for in data engineering?",
            "a": "SQL is used to query, join, filter, aggregate, and transform data stored in relational databases and many analytical platforms. It is also used to validate data and support reporting and warehouse transformations.",
            "keywords": ["query", "join", "aggregate", "transform", "validate"]
        },
        {
            "q": "What is the purpose of cloud storage?",
            "a": "Cloud storage stores data on remotely managed infrastructure. It offers scalable capacity, durability, access controls, and integration with data processing services. Permissions, encryption, retention, and cost should be managed carefully.",
            "keywords": ["scalable", "durability", "permissions", "encryption", "cost"]
        },
        {
            "q": "How do you ensure data quality?",
            "a": "I would define quality rules and validate completeness, uniqueness, accuracy, consistency, and acceptable ranges. I would monitor failures, log issues, alert the responsible team, and prevent invalid data from silently reaching downstream systems.",
            "keywords": ["completeness", "uniqueness", "consistency", "validation", "monitoring"]
        }
    ],

    "Java Developer": [
        {
            "q": "What are the main features of Java?",
            "a": "Java is an object-oriented, class-based language designed for portability. Its features include platform independence through bytecode and the JVM, automatic garbage collection, strong typing, exception handling, and a large standard library.",
            "keywords": ["object-oriented", "JVM", "platform independence", "garbage collection", "library"]
        },
        {
            "q": "What is the difference between JDK, JRE, and JVM?",
            "a": "The JVM executes Java bytecode. The JRE traditionally includes the JVM and runtime libraries needed to run Java applications. The JDK includes development tools such as the compiler, along with the components needed to develop and run Java programs.",
            "keywords": ["JVM", "runtime", "compiler", "development", "bytecode"]
        },
        {
            "q": "What are the four pillars of OOP?",
            "a": "The four pillars are encapsulation, abstraction, inheritance, and polymorphism. They help protect data, simplify interfaces, reuse code, and allow objects to behave differently through common interfaces.",
            "keywords": ["encapsulation", "abstraction", "inheritance", "polymorphism"]
        },
        {
            "q": "What is the difference between method overloading and overriding?",
            "a": "Overloading uses the same method name with different parameter lists in the same class or scope. Overriding occurs when a subclass provides its own implementation of an inherited instance method with a compatible signature.",
            "keywords": ["parameters", "same name", "subclass", "inherited", "implementation"]
        },
        {
            "q": "What is an exception in Java?",
            "a": "An exception is an event that interrupts normal program execution. Java handles exceptions using try, catch, finally, throw, and throws. Checked exceptions must be handled or declared, while unchecked exceptions include many programming errors.",
            "keywords": ["try", "catch", "finally", "checked", "unchecked"]
        },
        {
            "q": "Why is String immutable in Java?",
            "a": "A String object cannot be changed after it is created. Immutability helps with safe sharing, predictable behavior, string pooling, and security. Operations that appear to modify a String create a new String instead.",
            "keywords": ["immutable", "security", "string pool", "new object"]
        }
    ],

    "Web Developer": [
        {
            "q": "What is the difference between HTML, CSS, and JavaScript?",
            "a": "HTML defines the structure of a web page, CSS controls its presentation and layout, and JavaScript adds behavior and interactivity. Together they form the foundation of most front-end web applications.",
            "keywords": ["structure", "presentation", "interactivity", "HTML", "CSS", "JavaScript"]
        },
        {
            "q": "What happens when you enter a URL in a browser?",
            "a": "The browser identifies the destination, resolves its domain through DNS if needed, establishes a connection, sends an HTTP or HTTPS request, receives the response, and processes resources to render the page.",
            "keywords": ["DNS", "HTTP", "HTTPS", "request", "response", "render"]
        },
        {
            "q": "What is a REST API?",
            "a": "A REST API exposes resources through HTTP endpoints and commonly uses methods such as GET, POST, PUT, PATCH, and DELETE. It typically exchanges data in JSON and uses status codes to communicate the result of a request.",
            "keywords": ["HTTP", "GET", "POST", "resources", "JSON", "status codes"]
        },
        {
            "q": "What is responsive web design?",
            "a": "Responsive web design makes a website adapt to different screen sizes and devices. It uses flexible layouts, relative sizing, responsive images, and CSS media queries to provide a usable experience on phones, tablets, and desktops.",
            "keywords": ["screen sizes", "flexible", "media queries", "mobile", "layout"]
        },
        {
            "q": "What is the difference between front end and back end?",
            "a": "The front end is the part users interact with, including the interface and client-side behavior. The back end handles server-side logic, data processing, authentication, and database operations.",
            "keywords": ["interface", "client-side", "server-side", "database", "logic"]
        },
        {
            "q": "How would you make a website more secure?",
            "a": "I would validate inputs on the server, use parameterized database queries, protect credentials, enforce HTTPS, implement appropriate authentication and authorization, manage dependencies, and avoid exposing sensitive information in errors or logs.",
            "keywords": ["validation", "parameterized", "HTTPS", "authentication", "authorization"]
        }
    ],

    "Machine Learning Engineer": [
        {
            "q": "What is machine learning?",
            "a": "Machine learning is a field of AI in which systems learn patterns from data to make predictions or decisions. Common types include supervised learning, unsupervised learning, and reinforcement learning.",
            "keywords": ["patterns", "data", "supervised", "unsupervised", "predictions"]
        },
        {
            "q": "What is overfitting?",
            "a": "Overfitting occurs when a model learns the training data too closely, including noise, and performs poorly on unseen data. It can be reduced through validation, regularization, simpler models, cross-validation, and more representative data.",
            "keywords": ["training data", "unseen data", "validation", "regularization", "generalization"]
        },
        {
            "q": "What is feature engineering?",
            "a": "Feature engineering involves selecting, transforming, or creating input variables that help a model learn useful patterns. Examples include encoding categories, scaling numeric values, extracting date features, and creating meaningful ratios.",
            "keywords": ["selecting", "transforming", "creating", "variables", "encoding"]
        },
        {
            "q": "Why do we split data into training and testing sets?",
            "a": "The training set is used to fit the model, while the test set estimates how well it performs on unseen data. A separate validation set can be used for model selection and tuning without repeatedly using the test set.",
            "keywords": ["training", "testing", "unseen data", "validation", "evaluation"]
        },
        {
            "q": "What is the difference between classification and regression?",
            "a": "Classification predicts categories, such as spam or not spam. Regression predicts continuous numerical values, such as a house price or temperature. The evaluation metrics differ according to the task.",
            "keywords": ["categories", "continuous", "classification", "regression", "metrics"]
        },
        {
            "q": "Why might you choose Gradient Boosting?",
            "a": "Gradient Boosting builds an ensemble of weak learners, commonly decision trees, sequentially. Each new learner focuses on reducing the errors of the existing model. It can capture complex nonlinear relationships, although it requires careful tuning.",
            "keywords": ["ensemble", "decision trees", "sequentially", "errors", "nonlinear"]
        }
    ],

    "SQL Developer": [
        {
            "q": "What is SQL?",
            "a": "SQL stands for Structured Query Language. It is used to define, retrieve, insert, update, and delete data in relational databases, as well as manage database structures and permissions.",
            "keywords": ["Structured Query Language", "retrieve", "insert", "update", "relational"]
        },
        {
            "q": "What is a primary key?",
            "a": "A primary key uniquely identifies each row in a table. It must contain unique values and cannot be NULL. A table has one primary key constraint, which may consist of one or multiple columns.",
            "keywords": ["unique", "row", "NULL", "constraint", "columns"]
        },
        {
            "q": "What is a foreign key?",
            "a": "A foreign key references a candidate or primary key in another table, or sometimes the same table. It helps maintain referential integrity by ensuring that referenced records exist according to the defined constraints.",
            "keywords": ["references", "table", "referential integrity", "constraints"]
        },
        {
            "q": "What is normalization?",
            "a": "Normalization organizes relational tables to reduce unnecessary duplication and prevent insertion, update, and deletion anomalies. Normal forms such as 1NF, 2NF, and 3NF apply increasingly specific rules to table structure.",
            "keywords": ["duplication", "anomalies", "1NF", "2NF", "3NF"]
        },
        {
            "q": "What is the difference between WHERE and HAVING?",
            "a": "WHERE filters individual rows before grouping. HAVING filters groups after GROUP BY and is commonly used with aggregate conditions such as COUNT or SUM.",
            "keywords": ["rows", "groups", "GROUP BY", "aggregate"]
        },
        {
            "q": "What is an index in SQL?",
            "a": "An index is a data structure that can speed up data retrieval by helping the database locate rows efficiently. Indexes consume storage and can slow down inserts, updates, and deletes because they must also be maintained.",
            "keywords": ["retrieval", "data structure", "storage", "updates", "performance"]
        }
    ],

    "Software Tester": [
        {
            "q": "What is software testing?",
            "a": "Software testing evaluates whether an application behaves as expected and helps identify defects. It can involve manual or automated checks of functionality, performance, security, usability, and other quality attributes.",
            "keywords": ["expected", "defects", "manual", "automated", "quality"]
        },
        {
            "q": "What is the difference between verification and validation?",
            "a": "Verification checks whether the product is being built according to specifications. Validation checks whether the resulting product meets user needs and intended use.",
            "keywords": ["specifications", "built correctly", "user needs", "intended use"]
        },
        {
            "q": "What is regression testing?",
            "a": "Regression testing checks that recent code changes have not broken existing functionality. It often reuses automated test cases that cover important workflows and previously fixed defects.",
            "keywords": ["code changes", "existing functionality", "test cases", "defects"]
        },
        {
            "q": "What is Selenium?",
            "a": "Selenium is a suite of tools for automating web browsers. Selenium WebDriver lets tests interact with browser elements and verify application behavior using languages such as Java, Python, and C#.",
            "keywords": ["browser", "automation", "WebDriver", "Java", "web elements"]
        },
        {
            "q": "What is the difference between severity and priority?",
            "a": "Severity describes how much a defect affects the system. Priority describes how urgently it should be fixed. A defect can have high severity but a different priority depending on business needs and release timing.",
            "keywords": ["impact", "urgency", "defect", "business", "fix"]
        },
        {
            "q": "What is the purpose of TestNG?",
            "a": "TestNG is a Java testing framework. It supports annotations, test grouping, parameterization, dependencies, assertions, and test reports, making it useful for organizing and running automated tests.",
            "keywords": ["Java", "annotations", "assertions", "groups", "reports"]
        }
    ],

    "Android Developer": [
        {
            "q": "What is Android development?",
            "a": "Android development involves creating applications for Android devices. Developers use tools such as Android Studio and languages including Kotlin and Java to build interfaces, manage application logic, access device features, and test apps.",
            "keywords": ["Android Studio", "Kotlin", "Java", "interfaces", "testing"]
        },
        {
            "q": "What is Kotlin?",
            "a": "Kotlin is a statically typed programming language that runs on the JVM and is widely used for Android development. It offers concise syntax, null-safety features, extension functions, and interoperability with Java.",
            "keywords": ["statically typed", "JVM", "null-safety", "Android", "Java"]
        },
        {
            "q": "What is an Activity in Android?",
            "a": "An Activity is an Android application component that represents a screen or a focused interaction with the user. Its lifecycle includes callbacks such as onCreate, onStart, onResume, onPause, and onDestroy.",
            "keywords": ["screen", "component", "lifecycle", "onCreate", "onResume"]
        },
        {
            "q": "How can an Android app store data?",
            "a": "Android apps can store data using shared preferences or DataStore for small settings, files for documents, and Room or SQLite for structured local data. The appropriate choice depends on the type and volume of information.",
            "keywords": ["DataStore", "files", "Room", "SQLite", "local data"]
        },
        {
            "q": "How would you debug an Android application?",
            "a": "I would reproduce the problem, inspect Logcat messages, use breakpoints in Android Studio, check lifecycle and network behavior, and test on relevant devices or emulators. I would then fix the cause and retest the affected features.",
            "keywords": ["Logcat", "breakpoints", "Android Studio", "devices", "retest"]
        }
    ],

    "Data Scientist": [
        {
            "q": "What is the role of a data scientist?",
            "a": "A data scientist uses statistics, programming, and machine learning to analyze data, develop predictive models, and communicate findings. The work often includes defining a problem, preparing data, evaluating models, and presenting results.",
            "keywords": ["statistics", "programming", "machine learning", "models", "findings"]
        },
        {
            "q": "What is the difference between AI, machine learning, and deep learning?",
            "a": "AI is the broader field of creating systems that perform tasks associated with intelligence. Machine learning is a subset of AI that learns from data. Deep learning is a subset of machine learning based on multilayer neural networks.",
            "keywords": ["AI", "subset", "data", "neural networks", "machine learning"]
        },
        {
            "q": "What is a confusion matrix?",
            "a": "A confusion matrix compares predicted classes with actual classes. For binary classification it includes true positives, true negatives, false positives, and false negatives. These values are used to calculate metrics such as precision, recall, and accuracy.",
            "keywords": ["predicted", "actual", "true positives", "false positives", "precision", "recall"]
        },
        {
            "q": "What is the difference between precision and recall?",
            "a": "Precision is the proportion of predicted positives that are correct. Recall is the proportion of actual positives that the model successfully identifies. The preferred metric depends on the cost of false positives and false negatives.",
            "keywords": ["predicted positives", "actual positives", "false positives", "false negatives"]
        },
        {
            "q": "What is cross-validation?",
            "a": "Cross-validation evaluates a model across multiple splits of the training data. In k-fold cross-validation, the data is divided into k parts and each part is used for validation once. It gives a more stable estimate than a single split.",
            "keywords": ["splits", "k-fold", "training data", "validation", "estimate"]
        }
    ],

    "DevOps Engineer": [
        {
            "q": "What is DevOps?",
            "a": "DevOps combines development and operations practices to improve collaboration and deliver software more reliably. It often includes automation, continuous integration, continuous delivery, infrastructure as code, monitoring, and feedback.",
            "keywords": ["development", "operations", "automation", "CI/CD", "monitoring"]
        },
        {
            "q": "What is CI/CD?",
            "a": "Continuous Integration automatically builds and tests code changes frequently. Continuous Delivery keeps software ready for release, while Continuous Deployment automatically releases changes that pass the required checks.",
            "keywords": ["build", "tests", "integration", "delivery", "deployment"]
        },
        {
            "q": "What is Docker?",
            "a": "Docker is a platform for building and running applications in containers. Containers package an application and its dependencies into a consistent environment, helping reduce differences between development, testing, and deployment.",
            "keywords": ["containers", "dependencies", "consistent", "environment", "deployment"]
        },
        {
            "q": "What is the difference between a container and a virtual machine?",
            "a": "A virtual machine runs its own guest operating system on virtualized hardware. Containers generally share the host operating system kernel while isolating applications and their dependencies, making them typically lighter to start and run.",
            "keywords": ["guest operating system", "kernel", "isolation", "lighter", "virtualized"]
        },
        {
            "q": "How would you monitor a production application?",
            "a": "I would monitor availability, response times, error rates, resource usage, and important business metrics. Logs, metrics, traces, dashboards, and alerts help detect problems, investigate root causes, and measure recovery.",
            "keywords": ["availability", "logs", "metrics", "traces", "alerts", "root causes"]
        }
    ]
}

# --------------------------------------------------
# SESSION STATE
# --------------------------------------------------

DEFAULTS = {
    "page": "home",
    "resume_text": "",
    "resume_source": "",
    "selected_role": "",
    "question_index": 0,
    "answers": {},
    "show_model": {},
    "question_order": [],
    "session_started": False
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# --------------------------------------------------
# HELPER FUNCTIONS
# --------------------------------------------------

def extract_resume_text(uploaded_file):
    """Extract readable text from supported resume formats."""
    file_name = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    if file_name.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(file_bytes))
        pages = []

        for page in reader.pages:
            pages.append(page.extract_text() or "")

        return "\n".join(pages)

    elif file_name.endswith(".docx"):
        document = Document(io.BytesIO(file_bytes))
        paragraphs = [p.text for p in document.paragraphs]

        for table in document.tables:
            for row in table.rows:
                paragraphs.append(" | ".join(cell.text for cell in row.cells))

        return "\n".join(paragraphs)

    elif file_name.endswith(".pptx"):
        presentation = Presentation(io.BytesIO(file_bytes))
        text_parts = []

        for slide in presentation.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text_parts.append(shape.text)

        return "\n".join(text_parts)

    elif file_name.endswith(".txt"):
        return file_bytes.decode("utf-8", errors="ignore")

    else:
        raise ValueError(
            "Unsupported file type. Please upload a PDF, DOCX, PPTX, or TXT file."
        )


def get_resume_skills(resume_text):
    """Find a few recognizable technologies mentioned in the resume."""
    known_skills = [
        "Python", "Java", "SQL", "HTML", "CSS", "JavaScript",
        "Kotlin", "Flask", "Streamlit", "React", "Node.js",
        "Machine Learning", "Deep Learning", "Power BI", "Tableau",
        "Excel", "AWS", "Azure", "Docker", "Git", "GitHub",
        "Pandas", "NumPy", "Scikit-learn", "TensorFlow",
        "PyTorch", "Selenium", "TestNG", "MySQL", "MongoDB",
        "C++", "C", "Data Analysis", "Data Visualization"
    ]

    found = []
    resume_lower = resume_text.lower()

    for skill in known_skills:
        if skill.lower() in resume_lower:
            found.append(skill)

    return found[:12]


def get_personalized_questions(role, resume_text):
    """Return the question bank for the chosen role and optionally add a resume question."""
    questions = [item.copy() for item in QUESTION_BANK[role]]
    skills = get_resume_skills(resume_text)

    if skills:
        skill_list = ", ".join(skills[:5])

        personalized_question = {
            "q": f"Your resume mentions {skill_list}. How have you used these skills in a project?",
            "a": (
                "A strong answer should explain the project objective, the specific skills "
                "you used, your individual contribution, a challenge you faced, how you "
                "solved it, and the result. Be truthful about your own experience and "
                "distinguish what you implemented from what your team implemented."
            ),
            "keywords": [
                "project", "objective", "used", "contribution",
                "challenge", "solved", "result"
            ]
        }

        questions.insert(0, personalized_question)

    return questions


def score_answer(answer, question_data):
    """
    A simple keyword-coverage estimate.
    This is not an AI-based evaluation or a measure of factual correctness.
    """
    answer = answer.strip()

    if not answer:
        return 0, [], question_data["keywords"]

    answer_lower = answer.lower()
    keywords = question_data.get("keywords", [])

    matched = [
        keyword for keyword in keywords
        if keyword.lower() in answer_lower
    ]

    missing = [
        keyword for keyword in keywords
        if keyword.lower() not in answer_lower
    ]

    keyword_score = (len(matched) / max(len(keywords), 1)) * 70
    length_score = min(len(answer.split()) / 60, 1) * 30
    total_score = round(min(keyword_score + length_score, 100))

    return total_score, matched, missing


def get_score_feedback(score):
    if score >= 80:
        return (
            "Strong coverage",
            "Your answer includes many expected concepts. Check that it is accurate "
            "and support it with a practical example where possible."
        )
    elif score >= 55:
        return (
            "Good starting point",
            "Your answer covers some expected concepts. Add missing technical points "
            "and explain how they apply in practice."
        )
    elif score >= 30:
        return (
            "Needs more detail",
            "Try answering in a clear structure: definition, explanation, example, "
            "and conclusion."
        )
    else:
        return (
            "Keep practising",
            "Start with a simple definition, explain the main idea, and add a small "
            "example. You can review the sample answer for guidance."
        )


def reset_app():
    for key in list(st.session_state.keys()):
        del st.session_state[key]

    for key, value in DEFAULTS.items():
        st.session_state[key] = value


def render_brand():
    st.markdown("""
    <div class="brand-header">
        <div class="brand-logo">🎯</div>
        <div class="brand-name">
            <span class="brand-first">Interview</span><span class="brand-last">Preps</span>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_footer():
    st.markdown("""
    <div class="footer">
        InterviewPreps · Practise with purpose. Improve with every answer.
    </div>
    """, unsafe_allow_html=True)


def go_to_role_page():
    if not st.session_state.resume_text.strip():
        st.error("Please upload your resume or paste your resume text first.")
        return

    st.session_state.page = "role"
    st.rerun()


# --------------------------------------------------
# HEADER
# --------------------------------------------------

header_left, header_right = st.columns([5, 1])

with header_left:
    render_brand()

with header_right:
    if st.button("↻ Reset", key="reset_top", use_container_width=True):
        reset_app()
        st.rerun()


# --------------------------------------------------
# PAGE 1: HOME AND RESUME
# --------------------------------------------------

if st.session_state.page == "home":

    st.markdown("""
    <div class="hero">
        <div class="hero-tag">YOUR PERSONAL INTERVIEW PRACTICE SPACE</div>
        <h1>Prepare smarter.<br>Interview with confidence.</h1>
        <p>
            Practise for your next opportunity with questions tailored to your
            target role and resume. Write your answers, review sample responses,
            and discover which important concepts you could include.
        </p>
    </div>
    """, unsafe_allow_html=True)

    feature1, feature2, feature3 = st.columns(3)

    with feature1:
        st.markdown("""
        <div class="info-card">
            <div class="step-label">01 · PERSONALIZE</div>
            <h3>📄 Your resume</h3>
            <p>Upload your resume or paste its text to personalize your preparation.</p>
        </div>
        """, unsafe_allow_html=True)

    with feature2:
        st.markdown("""
        <div class="info-card">
            <div class="step-label">02 · SELECT</div>
            <h3>💼 Your target role</h3>
            <p>Choose from 11 technical roles and practise relevant interview questions.</p>
        </div>
        """, unsafe_allow_html=True)

    with feature3:
        st.markdown("""
        <div class="info-card">
            <div class="step-label">03 · IMPROVE</div>
            <h3>📈 Your answers</h3>
            <p>Review sample answers, check keyword coverage, and find areas to improve.</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("## Let's get started")
    st.markdown(
        "Choose **one** way to provide your resume. Your resume text is used "
        "within this Streamlit session to personalize the questions."
    )

    upload_tab, paste_tab = st.tabs(
        ["📎 Upload Resume", "✍️ Paste Resume Text"]
    )

    with upload_tab:
        st.markdown("### Upload your resume")
        st.caption("Supported formats: PDF, DOCX, PPTX, TXT")

        uploaded_file = st.file_uploader(
            "Choose your resume file",
            type=["pdf", "docx", "pptx", "txt"],
            key="resume_uploader",
            help="For scanned PDFs, text extraction may not work unless the PDF contains selectable text."
        )

        if uploaded_file is not None:
            if st.button(
                "Extract Resume Text",
                type="primary",
                use_container_width=True,
                key="extract_resume"
            ):
                try:
                    extracted_text = extract_resume_text(uploaded_file)

                    if not extracted_text.strip():
                        st.error(
                            "No readable text was found. If this is a scanned PDF, "
                            "try a text-based PDF or paste your resume text instead."
                        )
                    else:
                        st.session_state.resume_text = extracted_text
                        st.session_state.resume_source = uploaded_file.name
                        st.success("Resume text extracted successfully!")
                        st.rerun()

                except Exception as error:
                    st.error(f"Could not read this file: {error}")

    with paste_tab:
        st.markdown("### Paste your resume")
        pasted_resume = st.text_area(
            "Resume content",
            value=(
                st.session_state.resume_text
                if st.session_state.resume_source == "Pasted text"
                else ""
            ),
            height=250,
            placeholder=(
                "Paste your resume here...\n\n"
                "Include your summary, technical skills, projects, education, "
                "internships, and certifications."
            ),
            key="resume_paste_box"
        )

        if st.button(
            "Use Pasted Resume",
            type="primary",
            use_container_width=True,
            key="use_pasted_resume"
        ):
            if len(pasted_resume.strip()) < 30:
                st.error("Please paste a little more resume content before continuing.")
            else:
                st.session_state.resume_text = pasted_resume.strip()
                st.session_state.resume_source = "Pasted text"
                st.success("Resume text saved for this practice session!")
                st.rerun()

    if st.session_state.resume_text.strip():
        st.markdown("---")
        st.markdown("### Resume ready")

        resume_words = len(st.session_state.resume_text.split())
        resume_skills = get_resume_skills(st.session_state.resume_text)

        ready_col1, ready_col2 = st.columns(2)

        with ready_col1:
            st.markdown(f"""
            <div class="info-card">
                <div class="step-label">RESUME SOURCE</div>
                <h3>📄 {st.session_state.resume_source or "Resume text"}</h3>
                <p>{resume_words} words extracted or provided</p>
            </div>
            """, unsafe_allow_html=True)

        with ready_col2:
            skill_text = (
                ", ".join(resume_skills)
                if resume_skills
                else "No common technology keywords detected yet"
            )
            st.markdown(f"""
            <div class="info-card">
                <div class="step-label">DETECTED SKILLS</div>
                <h3>🧩 Resume overview</h3>
                <p>{skill_text}</p>
            </div>
            """, unsafe_allow_html=True)

        if st.button(
            "Continue to Choose Your Role →",
            type="primary",
            use_container_width=True,
            key="continue_to_role"
        ):
            go_to_role_page()

    render_footer()


# --------------------------------------------------
# PAGE 2: ROLE SELECTION
# --------------------------------------------------

elif st.session_state.page == "role":

    st.markdown("""
    <div class="hero">
        <div class="hero-tag">STEP 2 OF 3</div>
        <h1>Choose your target role.</h1>
        <p>
            Select the role you're preparing for. InterviewPreps will load
            questions relevant to that job profile and look for recognizable
            skills in your resume.
        </p>
    </div>
    """, unsafe_allow_html=True)

    if st.button("← Back to Resume", key="back_to_resume"):
        st.session_state.page = "home"
        st.rerun()

    st.markdown("### Which role are you preparing for?")

    roles = list(QUESTION_BANK.keys())

    role_descriptions = {
        "Python Developer": "Python programming, functions, data structures, and debugging",
        "Data Analyst": "Data cleaning, SQL, statistics, and visualization",
        "Cloud Data Engineer": "ETL, data pipelines, cloud storage, and data quality",
        "Java Developer": "Core Java, OOP, exceptions, and JVM fundamentals",
        "Web Developer": "HTML, CSS, JavaScript, APIs, and web security",
        "Machine Learning Engineer": "ML concepts, feature engineering, and evaluation",
        "SQL Developer": "Queries, joins, normalization, and database performance",
        "Software Tester": "Testing fundamentals, Selenium, TestNG, and defects",
        "Android Developer": "Kotlin, Android components, storage, and debugging",
        "Data Scientist": "Statistics, machine learning, and model evaluation",
        "DevOps Engineer": "CI/CD, Docker, deployment, and monitoring"
    }

    selected_role = st.selectbox(
        "Select a job role",
        options=roles,
        index=(
            roles.index(st.session_state.selected_role)
            if st.session_state.selected_role in roles
            else 0
        ),
        key="role_selector"
    )

    st.markdown(f"""
    <div class="info-card">
        <div class="step-label">ROLE OVERVIEW</div>
        <h3>💼 {selected_role}</h3>
        <p>{role_descriptions[selected_role]}</p>
        <p><strong>{len(QUESTION_BANK[selected_role])} core questions</strong> available,
        plus an additional resume-based question when recognizable skills are detected.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Your practice settings")

    practice_mode = st.radio(
        "How would you like to practise?",
        ["One question at a time", "Review all questions"],
        horizontal=True,
        key="practice_mode"
    )

    if st.button(
        "Start Interview Practice →",
        type="primary",
        use_container_width=True,
        key="start_interview"
    ):
        st.session_state.selected_role = selected_role
        st.session_state.question_index = 0
        st.session_state.answers = {}
        st.session_state.show_model = {}
        st.session_state.question_order = get_personalized_questions(
            selected_role,
            st.session_state.resume_text
        )
        st.session_state.session_started = True
        st.session_state.page = "interview"
        st.session_state.practice_mode_selected = practice_mode
        st.rerun()

    render_footer()


# --------------------------------------------------
# PAGE 3: INTERVIEW PRACTICE
# --------------------------------------------------

elif st.session_state.page == "interview":

    role = st.session_state.selected_role
    questions = st.session_state.question_order

    if not questions:
        st.error("No questions were loaded. Please select your role again.")
        if st.button("Back to role selection"):
            st.session_state.page = "role"
            st.rerun()
        st.stop()

    total_questions = len(questions)
    mode = st.session_state.get(
        "practice_mode_selected",
        "One question at a time"
    )

    st.markdown(f"""
    <div class="hero">
        <div class="hero-tag">STEP 3 OF 3 · INTERVIEW PRACTICE</div>
        <h1>Let's practise, one answer at a time.</h1>
        <p>
            Target role: <strong>{role}</strong><br>
            Take your time. Write your own answer before looking at the sample response.
        </p>
    </div>
    """, unsafe_allow_html=True)

    top_left, top_right = st.columns([4, 1])

    with top_left:
        st.markdown(f"### Your interview: {role}")

    with top_right:
        if st.button("End Practice", key="end_practice"):
            st.session_state.page = "report"
            st.rerun()

    if mode == "One question at a time":

        current_index = st.session_state.question_index
        current_index = max(0, min(current_index, total_questions - 1))
        question_data = questions[current_index]
        question_key = f"answer_{current_index}"

        st.progress(
            (current_index + 1) / total_questions,
            text=f"Question {current_index + 1} of {total_questions}"
        )

        st.markdown(f"""
        <div class="info-card">
            <div class="step-label">QUESTION {current_index + 1}</div>
            <h3>{question_data["q"]}</h3>
            <p>Think about the concept, explain it clearly, and include an example when possible.</p>
        </div>
        """, unsafe_allow_html=True)

        answer_text = st.text_area(
            "Your answer",
            value=st.session_state.answers.get(question_key, ""),
            height=200,
            placeholder=(
                "Write your answer here...\n\n"
                "Tip: Start with a definition, explain the key idea, and give an example."
            ),
            key=f"answer_box_{current_index}"
        )

        st.session_state.answers[question_key] = answer_text

        action_col1, action_col2 = st.columns(2)

        with action_col1:
            if st.button(
                "Check My Answer",
                type="primary",
                use_container_width=True,
                key=f"check_answer_{current_index}"
            ):
                if not answer_text.strip():
                    st.warning("Write an answer first so you can review it.")
                else:
                    score, matched, missing = score_answer(answer_text, question_data)
                    feedback_title, feedback_text = get_score_feedback(score)

                    st.markdown("---")
                    st.markdown("### Your answer review")

                    metric_col1, metric_col2 = st.columns(2)

                    with metric_col1:
                        st.metric("Keyword coverage estimate", f"{score}/100")

                    with metric_col2:
                        st.metric("Words in your answer", len(answer_text.split()))

                    st.progress(score / 100)

                    st.info(
                        f"**{feedback_title}:** {feedback_text}"
                    )

                    st.caption(
                        "This score is a simple heuristic based on keyword coverage "
                        "and answer length. It does not verify factual correctness."
                    )

                    if matched:
                        st.markdown("**Concepts detected in your answer**")
                        st.write(", ".join(matched))

                    if missing:
                        st.markdown("**Concepts you may want to discuss**")
                        st.write(", ".join(missing))

                    st.markdown("**Suggested answer structure**")
                    st.write(
                        "1. Define the concept.\n"
                        "2. Explain how it works.\n"
                        "3. Give an example or practical use.\n"
                        "4. Mention a benefit, limitation, or best practice."
                    )

        with action_col2:
            if st.button(
                "Show Sample Answer",
                use_container_width=True,
                key=f"show_answer_{current_index}"
            ):
                st.session_state.show_model[question_key] = True

        if st.session_state.show_model.get(question_key, False):
            st.markdown("---")
            st.markdown("### Sample answer")

            st.markdown(f"""
            <div class="info-card">
                <p>{question_data["a"]}</p>
            </div>
            """, unsafe_allow_html=True)

            st.caption(
                "Use this as a learning reference. Adapt the answer to your own "
                "understanding and real experience instead of memorizing it word for word."
            )

        st.markdown("---")

        nav_left, nav_mid, nav_right = st.columns([1, 2, 1])

        with nav_left:
            if st.button(
                "← Previous",
                disabled=current_index == 0,
                use_container_width=True,
                key="previous_question"
            ):
                st.session_state.question_index -= 1
                st.rerun()

        with nav_mid:
            if st.button(
                "View Progress Report",
                use_container_width=True,
                key="view_report"
            ):
                st.session_state.page = "report"
                st.rerun()

        with nav_right:
            if current_index < total_questions - 1:
                if st.button(
                    "Next →",
                    type="primary",
                    use_container_width=True,
                    key="next_question"
                ):
                    st.session_state.question_index += 1
                    st.rerun()
            else:
                if st.button(
                    "Finish →",
                    type="primary",
                    use_container_width=True,
                    key="finish_interview"
                ):
                    st.session_state.page = "report"
                    st.rerun()

    else:
        st.progress(1.0, text=f"All {total_questions} questions")

        st.info(
            "Review mode lets you work through all questions on one page. "
            "Your answers are kept while you remain in this session."
        )

        for i, question_data in enumerate(questions):
            question_key = f"answer_{i}"

            with st.expander(
                f"Question {i + 1}: {question_data['q']}",
                expanded=(i == 0)
            ):
                answer_text = st.text_area(
                    "Your answer",
                    value=st.session_state.answers.get(question_key, ""),
                    height=150,
                    placeholder="Write your answer here...",
                    key=f"review_answer_{i}"
                )

                st.session_state.answers[question_key] = answer_text

                button_col1, button_col2 = st.columns(2)

                with button_col1:
                    if st.button(
                        "Review Answer",
                        key=f"review_check_{i}",
                        use_container_width=True
                    ):
                        if not answer_text.strip():
                            st.warning("Write an answer first.")
                        else:
                            score, matched, missing = score_answer(
                                answer_text,
                                question_data
                            )
                            feedback_title, feedback_text = get_score_feedback(score)

                            st.metric("Keyword coverage estimate", f"{score}/100")
                            st.progress(score / 100)
                            st.info(f"**{feedback_title}:** {feedback_text}")

                            if missing:
                                st.write("Concepts to consider:", ", ".join(missing))

                            st.caption(
                                "This is a keyword-based estimate, not an AI assessment."
                            )

                with button_col2:
                    if st.button(
                        "Show Sample Answer",
                        key=f"review_model_{i}",
                        use_container_width=True
                    ):
                        st.session_state.show_model[question_key] = True

                if st.session_state.show_model.get(question_key, False):
                    st.markdown("**Sample answer**")
                    st.write(question_data["a"])

        if st.button(
            "Generate Progress Report →",
            type="primary",
            use_container_width=True,
            key="review_report"
        ):
            st.session_state.page = "report"
            st.rerun()

    render_footer()


# --------------------------------------------------
# PAGE 4: PROGRESS REPORT
# --------------------------------------------------

elif st.session_state.page == "report":

    role = st.session_state.selected_role
    questions = st.session_state.question_order
    total_questions = len(questions)

    st.markdown("""
    <div class="hero">
        <div class="hero-tag">YOUR PRACTICE SUMMARY</div>
        <h1>Every practice session is progress.</h1>
        <p>
            Review your completed answers, identify topics to revisit, and
            decide what to practise next.
        </p>
    </div>
    """, unsafe_allow_html=True)

    answered_count = 0
    scores = []
    report_rows = []

    for i, question_data in enumerate(questions):
        question_key = f"answer_{i}"
        answer_text = st.session_state.answers.get(question_key, "").strip()

        if answer_text:
            answered_count += 1
            score, matched, missing = score_answer(answer_text, question_data)
            scores.append(score)

            report_rows.append({
                "number": i + 1,
                "question": question_data["q"],
                "score": score,
                "missing": missing,
                "answered": True
            })
        else:
            report_rows.append({
                "number": i + 1,
                "question": question_data["q"],
                "score": 0,
                "missing": question_data.get("keywords", []),
                "answered": False
            })

    average_score = round(sum(scores) / len(scores)) if scores else 0
    completion = round((answered_count / total_questions) * 100) if total_questions else 0

    st.markdown(f"### Target role: {role}")

    metric1, metric2, metric3 = st.columns(3)

    with metric1:
        st.metric("Questions answered", f"{answered_count}/{total_questions}")

    with metric2:
        st.metric("Completion", f"{completion}%")

    with metric3:
        st.metric(
            "Average keyword estimate",
            f"{average_score}/100" if scores else "N/A"
        )

    st.progress(completion / 100 if total_questions else 0)

    st.markdown("---")
    st.markdown("### Question-by-question review")

    for item in report_rows:
        with st.expander(
            f"Q{item['number']}. {item['question']}"
        ):
            if not item["answered"]:
                st.warning("Not answered yet.")
                st.write("Suggested concepts:", ", ".join(item["missing"]))
            else:
                st.metric("Keyword coverage estimate", f"{item['score']}/100")
                st.progress(item["score"] / 100)

                if item["missing"]:
                    st.write(
                        "**Concepts to review:**",
                        ", ".join(item["missing"])
                    )
                else:
                    st.success("All listed keywords were detected in your answer.")

                if st.button(
                    "Show sample answer",
                    key=f"report_model_{item['number']}"
                ):
                    st.session_state.show_model[
                        f"report_{item['number']}"
                    ] = True

                if st.session_state.show_model.get(
                    f"report_{item['number']}", False
                ):
                    st.markdown("**Sample answer**")
                    st.write(questions[item["number"] - 1]["a"])

    st.markdown("---")
    st.markdown("### What should you do next?")

    if answered_count == 0:
        st.info(
            "You haven't submitted any answers yet. Return to the interview, "
            "try a question, and use the sample answer to guide your practice."
        )
    elif average_score < 55:
        st.info(
            "Review the sample answers and practise explaining each concept "
            "in your own words. Add definitions, key technical points, and examples."
        )
    else:
        st.info(
            "Good progress. Continue practising, verify technical accuracy, "
            "and use real project examples to make your answers stronger."
        )

    st.caption(
        "Important: scores are generated using keyword matching and answer length. "
        "They are not a measure of factual correctness, communication quality, or "
        "your actual interview performance."
    )

    left_action, right_action = st.columns(2)

    with left_action:
        if st.button(
            "← Return to Questions",
            use_container_width=True,
            key="return_questions"
        ):
            st.session_state.page = "interview"
            st.rerun()

    with right_action:
        if st.button(
            "Start a New Practice Session",
            type="primary",
            use_container_width=True,
            key="new_practice"
        ):
            reset_app()
            st.rerun()

    render_footer()