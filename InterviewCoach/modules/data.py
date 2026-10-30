import streamlit as st


ROLES = [
    "Python Developer",
    "Java Developer",
    "Data Analyst",
    "AI/ML Engineer",
    "Full Stack Developer",
    "Software Developer"
]


ROLE_SKILLS = {

    "Python Developer": [
        "Python",
        "OOP",
        "Functions",
        "Exception Handling",
        "SQL",
        "Git",
        "APIs",
        "Testing"
    ],

    "Java Developer": [
        "Java",
        "OOP",
        "Collections",
        "Exception Handling",
        "SQL",
        "JDBC",
        "Spring",
        "Git"
    ],

    "Data Analyst": [
        "Python",
        "SQL",
        "Excel",
        "Pandas",
        "Data Cleaning",
        "Statistics",
        "Visualization",
        "Power BI"
    ],

    "AI/ML Engineer": [
        "Python",
        "NumPy",
        "Pandas",
        "Statistics",
        "Machine Learning",
        "Scikit-learn",
        "Deep Learning",
        "SQL"
    ],

    "Full Stack Developer": [
        "HTML",
        "CSS",
        "JavaScript",
        "React",
        "REST API",
        "Node.js",
        "SQL",
        "Git"
    ],

    "Software Developer": [
        "Programming",
        "OOP",
        "DSA",
        "DBMS",
        "SQL",
        "Git",
        "Testing",
        "Problem Solving"
    ]
}


QUESTION_BANK = {

    "Python Developer": {

        "Technical": [

            (
                "What is OOP in Python?",
                [
                    "class",
                    "object",
                    "inheritance",
                    "encapsulation",
                    "polymorphism"
                ]
            ),

            (
                "Explain the difference between a list and a tuple.",
                [
                    "list",
                    "tuple",
                    "mutable",
                    "immutable"
                ]
            ),

            (
                "What are Python decorators?",
                [
                    "decorator",
                    "function",
                    "wrapper",
                    "@"
                ]
            ),

            (
                "How does exception handling work in Python?",
                [
                    "try",
                    "except",
                    "finally",
                    "raise"
                ]
            ),

            (
                "What is a dictionary in Python?",
                [
                    "key",
                    "value",
                    "hash",
                    "lookup"
                ]
            ),

            (
                "What is the difference between == and is?",
                [
                    "value",
                    "identity",
                    "object"
                ]
            )
        ],

        "HR": [

            (
                "Tell me about yourself.",
                [
                    "education",
                    "skills",
                    "project",
                    "goal"
                ]
            ),

            (
                "Why do you want to become a Python Developer?",
                [
                    "Python",
                    "programming",
                    "project",
                    "career"
                ]
            )
        ],

        "Behavioral": [

            (
                "Tell me about a time you solved a difficult problem.",
                [
                    "situation",
                    "problem",
                    "action",
                    "result"
                ]
            ),

            (
                "How do you handle learning a new technology?",
                [
                    "learn",
                    "practice",
                    "documentation",
                    "project"
                ]
            )
        ]
    },


    "Java Developer": {

        "Technical": [

            (
                "Explain inheritance in Java.",
                [
                    "class",
                    "parent",
                    "child",
                    "extends",
                    "reuse"
                ]
            ),

            (
                "What is polymorphism?",
                [
                    "overloading",
                    "overriding",
                    "runtime",
                    "compile"
                ]
            ),

            (
                "What is the difference between JDK, JRE and JVM?",
                [
                    "JDK",
                    "JRE",
                    "JVM",
                    "compile",
                    "runtime"
                ]
            ),

            (
                "Explain encapsulation in Java.",
                [
                    "private",
                    "getter",
                    "setter",
                    "data"
                ]
            )
        ],

        "HR": [

            (
                "Tell me about yourself.",
                [
                    "education",
                    "skills",
                    "project",
                    "goal"
                ]
            ),

            (
                "Why do you want to work as a Java Developer?",
                [
                    "Java",
                    "programming",
                    "project",
                    "career"
                ]
            )
        ],

        "Behavioral": [

            (
                "Describe a challenging project you worked on.",
                [
                    "problem",
                    "action",
                    "result",
                    "team"
                ]
            )
        ]
    },


    "Data Analyst": {

        "Technical": [

            (
                "What is data cleaning?",
                [
                    "missing",
                    "duplicate",
                    "outlier",
                    "data"
                ]
            ),

            (
                "Explain normalization in databases.",
                [
                    "normalization",
                    "redundancy",
                    "1NF",
                    "2NF",
                    "3NF"
                ]
            ),

            (
                "What is a primary key?",
                [
                    "unique",
                    "null",
                    "row",
                    "identifier"
                ]
            ),

            (
                "What is the difference between INNER JOIN and LEFT JOIN?",
                [
                    "matching",
                    "left",
                    "rows",
                    "join"
                ]
            ),

            (
                "What is the purpose of GROUP BY in SQL?",
                [
                    "aggregate",
                    "group",
                    "count",
                    "sum"
                ]
            )
        ],

        "HR": [

            (
                "Why do you want to become a Data Analyst?",
                [
                    "data",
                    "analysis",
                    "insight",
                    "problem"
                ]
            )
        ],

        "Behavioral": [

            (
                "Tell me about a time you used data to make a decision.",
                [
                    "data",
                    "analysis",
                    "decision",
                    "result"
                ]
            )
        ]
    },


    "AI/ML Engineer": {

        "Technical": [

            (
                "What is supervised learning?",
                [
                    "label",
                    "training",
                    "classification",
                    "regression"
                ]
            ),

            (
                "Explain overfitting.",
                [
                    "training",
                    "test",
                    "generalization",
                    "regularization"
                ]
            ),

            (
                "What is a neural network?",
                [
                    "neuron",
                    "layer",
                    "weights",
                    "activation"
                ]
            ),

            (
                "What is the difference between AI and ML?",
                [
                    "AI",
                    "ML",
                    "learning",
                    "system"
                ]
            )
        ],

        "HR": [

            (
                "Tell me about yourself.",
                [
                    "education",
                    "skills",
                    "project",
                    "goal"
                ]
            )
        ],

        "Behavioral": [

            (
                "Describe a machine learning problem you solved.",
                [
                    "dataset",
                    "model",
                    "evaluation",
                    "result"
                ]
            )
        ]
    },


    "Full Stack Developer": {

        "Technical": [

            (
                "What is a REST API?",
                [
                    "HTTP",
                    "GET",
                    "POST",
                    "PUT",
                    "DELETE",
                    "resource"
                ]
            ),

            (
                "Explain frontend vs backend.",
                [
                    "UI",
                    "server",
                    "database",
                    "client"
                ]
            ),

            (
                "What is authentication?",
                [
                    "identity",
                    "login",
                    "token",
                    "password"
                ]
            ),

            (
                "What is a database?",
                [
                    "store",
                    "table",
                    "query",
                    "data"
                ]
            )
        ],

        "HR": [

            (
                "Tell me about yourself.",
                [
                    "education",
                    "skills",
                    "project",
                    "goal"
                ]
            )
        ],

        "Behavioral": [

            (
                "Describe a challenging web project.",
                [
                    "problem",
                    "frontend",
                    "backend",
                    "solution"
                ]
            )
        ]
    },


    "Software Developer": {

        "Technical": [

            (
                "What is Object Oriented Programming?",
                [
                    "class",
                    "object",
                    "inheritance",
                    "polymorphism"
                ]
            ),

            (
                "What is time complexity?",
                [
                    "input",
                    "growth",
                    "complexity",
                    "Big O"
                ]
            ),

            (
                "What is a database primary key?",
                [
                    "unique",
                    "identifier",
                    "row"
                ]
            ),

            (
                "Why is version control important?",
                [
                    "Git",
                    "history",
                    "branch",
                    "collaboration"
                ]
            )
        ],

        "HR": [

            (
                "Tell me about yourself.",
                [
                    "education",
                    "skills",
                    "project",
                    "goal"
                ]
            ),

            (
                "What are your strengths?",
                [
                    "skill",
                    "example",
                    "learning",
                    "problem"
                ]
            ),

            (
                "Why should we hire you?",
                [
                    "skills",
                    "value",
                    "learning",
                    "team"
                ]
            )
        ],

        "Behavioral": [

            (
                "Describe a challenging project.",
                [
                    "problem",
                    "action",
                    "result",
                    "team"
                ]
            )
        ]
    }
}


RESOURCES = [

    (
        "Python Fundamentals",
        "Revise variables, functions, collections, OOP and exception handling."
    ),

    (
        "SQL Practice",
        "Practice SELECT, JOIN, GROUP BY, subqueries and aggregate functions."
    ),

    (
        "DBMS",
        "Review normalization, keys, transactions, indexing and ACID properties."
    ),

    (
        "DSA",
        "Practice arrays, strings, stacks, queues, trees, hashing and complexity."
    ),

    (
        "HR Preparation",
        "Prepare concise STAR-format stories for projects and challenges."
    ),

    (
        "Communication",
        "Practice explaining technical concepts using definition, example and use case."
    )
]


def initialize_state():

    defaults = {

        "candidate_name": "Candidate",

        "target_role": "Python Developer",

        "resume_text": "",

        "history": [],

        "question": None,

        "answer": "",

        "question_category": "Technical",

        "difficulty": "Intermediate",

        "last_result": None
    }

    for key, value in defaults.items():

        if key not in st.session_state:

            st.session_state[key] = value