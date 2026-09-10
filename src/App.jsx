import React, { useState, useMemo } from "react";
import { Lock, Check, ArrowLeft, ShieldCheck, CheckCircle2, XCircle } from "lucide-react";
import * as XLSX from "xlsx";

// ---------------------------------------------------------------------------
// DATA — generated from screening-form-data.xlsx via build_config.py.
// Categories are NOT hardcoded: each package's `sections` object and each
// brand's `filterOverrides` carry whatever categories the spreadsheet used,
// in the order rows appeared. The component derives category lists at
// render time instead of assuming a fixed set.
// ---------------------------------------------------------------------------

const PACKAGES = [
  {
    "id": "reach",
    "name": "Reach",
    "blurb": "Widest reach of our audience",
    "sections": {}
  },
  {
    "id": "targeted",
    "name": "Targeted",
    "blurb": "Targeted focus on your core audience",
    "sections": {}
  },
  {
    "id": "precision",
    "name": "Precision",
    "blurb": "Precision leads to target just who you want",
    "sections": {}
  }
];

const BRANDS = [
  {
    "id": "BPRO",
    "name": "BenefitsPRO",
    "isCustom": null,
    "logoUrl": "https://www.benefitspro.com/_nuxt/img/bpro-blue.0a30307.png",
    "subAudiences": [
      "Benefits Brokers",
      "Human Resources",
      "Brokers & HR"
    ],
    "subAudienceCombos": {
      "Brokers & HR": [
        "Benefits Brokers",
        "Human Resources"
      ]
    },
    "accountReps": [
      {
        "name": "Alicia Robledo",
        "email": "alicia.robledo@arc-network.com"
      },
      {
        "name": "Kelly Davisson",
        "email": "kelly.davisson@arc-network.com"
      }
    ],
    "customQuestions": {
      "targeted": [
        {
          "id": "group-size",
          "question": "Do you or your organization work with groups of the following size? (must select at least two options)",
          "type": "optional",
          "subAudience": "Benefits Brokers",
          "standard": [
            "50 lives or fewer"
          ],
          "optional": [
            "51 - 100 lives",
            "101 - 500 lives ",
            "501 - 1,000 lives ",
            "1,001 - 5,000 lives ",
            "More than 5,000 lives "
          ]
        }
      ],
      "precision": [
        {
          "id": "group-size",
          "question": "Do you or your organization work with groups of the following size? (must select at least two options)",
          "type": "optional",
          "subAudience": "Benefits Brokers",
          "standard": [
            "50 lives or fewer"
          ],
          "optional": [
            "51 - 100 lives",
            "101 - 500 lives ",
            "501 - 1,000 lives ",
            "1,001 - 5,000 lives ",
            "More than 5,000 lives "
          ]
        },
        {
          "id": "self-funded",
          "question": "Do you work with self funded?",
          "type": "optional",
          "subAudience": "Benefits Brokers",
          "standard": [
            "Yes (accepted)",
            "No (not accepted)"
          ],
          "optional": []
        }
      ]
    },
    "filterOverrides": {
      "precision": [
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Advertising/Marketing",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Customer Support - Client Services/Account Management",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Aerospace/Aviation",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Agriculture",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Automotive",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Communications",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Employee Benefits Broker/Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Computers and Technology",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Development",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Diversity Management/EEO/Compliance",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Construction",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Education",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - HR Systems Administration",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Owner",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Learning/Training",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Finance",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "InsuranceSenior Management",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Government",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Organizational Development",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Healthcare/Medical",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Agent/Broker",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Recruiting",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Insurance",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Legal",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Manufacturing",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Talent Management",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Media",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Wellness",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Non-Profit/Organizations",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Real Estate",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Retail and Consumer Goods",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Service Industry",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Telecommunications",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Transportation and Logistics",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Travel/Hospitality/Entertainment",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Utility/Energy",
          "type": "optional"
        }
      ],
      "reach": [
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Customer Support - Client Services/Account Management",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Advertising/Marketing",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Aerospace/Aviation",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Agriculture",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Automotive",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Communications",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Employee Benefits Broker/Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Computers and Technology",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Development",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Diversity Management/EEO/Compliance",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Construction",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Education",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - HR Systems Administration",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Finance",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Learning/Training",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "InsuranceSenior Management",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Government",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Organizational Development",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Contractor",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Contractor",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Healthcare/Medical",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Agent/Broker",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Insurance",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Recruiting",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Legal",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Talent Management",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Manufacturing",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Media",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Wellness",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Non-Profit/Organizations",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Real Estate",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Retail and Consumer Goods",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Service Industry",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Telecommunications",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Transportation and Logistics",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Travel/Hospitality/Entertainment",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Utility/Energy",
          "type": "standard"
        }
      ],
      "targeted": [
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Advertising/Marketing",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Customer Support - Client Services/Account Management",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Aerospace/Aviation",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Agriculture",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Automotive",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Communications",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Employee Benefits Broker/Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Computers and Technology",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Development",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Diversity Management/EEO/Compliance",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Construction",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - HR Systems Administration",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Education",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Learning/Training",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Finance",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "InsuranceSenior Management",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Organizational Development",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Government",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Healthcare/Medical",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Agent/Broker",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Insurance",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Recruiting",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Legal",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Talent Management",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Manufacturing",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Media",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Wellness",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Non-Profit/Organizations",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Real Estate",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Retail and Consumer Goods",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Service Industry",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Telecommunications",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Transportation and Logistics",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Travel/Hospitality/Entertainment",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Utility/Energy",
          "type": "optional"
        }
      ]
    }
  },
  {
    "id": "TA",
    "name": "ThinkAdvisor",
    "isCustom": null,
    "logoUrl": "https://www.thinkadvisor.com/_nuxt/img/ALM_ThinkAdvisor.84d9263.png",
    "subAudiences": null,
    "subAudienceCombos": {},
    "accountReps": [
      {
        "name": "Archer Montague",
        "email": "archer.montague@arc-network.com"
      },
      {
        "name": "David Spindler",
        "email": "david.spindler@arc-network.com"
      }
    ],
    "customQuestions": {
      "targeted": [
        {
          "id": "firm-aum",
          "question": "What is your firm\u2019s assets under management:   ",
          "type": "optional",
          "subAudience": null,
          "standard": [
            "$1B or more (must accept)  "
          ],
          "optional": [
            "Less than $40M ",
            "$400M - $499M",
            "$500M - $749M ",
            "$750M - $999M "
          ]
        }
      ],
      "precision": [
        {
          "id": "ria",
          "question": "Are you an RIA?",
          "type": "required",
          "subAudience": null,
          "standard": [
            "Yes (accepted)",
            "No (not accepted)"
          ],
          "optional": []
        },
        {
          "id": "firm-aum",
          "question": "What is your firm\u2019s assets under management:   ",
          "type": "optional",
          "subAudience": null,
          "standard": [
            "$1B or more (must accept)  "
          ],
          "optional": [
            "Less than $40M ",
            "$400M - $499M",
            "$500M - $749M ",
            "$750M - $999M "
          ]
        }
      ]
    },
    "filterOverrides": {
      "precision": [
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial (General)",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Owner/Co-owner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Product Sales/Brokerage",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Registered Investment Advisor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Registered Rep",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Securities Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Actuarial Analysis",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Employee Benefits Broker/Consultant",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Financial Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Risk Management/Compliance",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Senior Management",
          "type": "optional"
        }
      ],
      "reach": [
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial (General)",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Product Sales/Brokerage",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Registered Investment Advisor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Registered Rep",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Securities Analysis/Research",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Actuarial Analysis",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Employee Benefits Broker/Consultant",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Financial Analysis/Research",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Risk Management/Compliance",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Senior Management",
          "type": "standard"
        }
      ],
      "targeted": [
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial (General)",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Product Sales/Brokerage",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Registered Investment Advisor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Registered Rep",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Securities Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Actuarial Analysis",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Employee Benefits Broker/Consultant",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Financial Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Risk Management/Compliance",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Insurance - Senior Management",
          "type": "optional"
        }
      ]
    }
  },
  {
    "id": "PC360",
    "name": "PropertyCasualty360",
    "isCustom": null,
    "logoUrl": "https://www.propertycasualty360.com/_nuxt/img/pc360-blue.013d9ad.png",
    "subAudiences": [
      "Entire PC360 Audience",
      "Agents & Brokers",
      "Carriers",
      "Claims"
    ],
    "subAudienceCombos": {},
    "accountReps": [
      {
        "name": "Susan Gould",
        "email": "Susan.Gould@arc-network.com"
      },
      {
        "name": "Kelly Davisson",
        "email": "kelly.davisson@arc-network.com"
      }
    ],
    "customQuestions": {},
    "filterOverrides": {
      "reach": [
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Automotive",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Automotive",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Agriculture",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Automotive",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Owner/Co-owner",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Property & Casualty",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Property & Casualty",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Managing Partner",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Property & Casualty",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Proprietor",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Contractor",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Agent/Broker",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Contractor",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Business Development",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Field Service",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - General Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Owner/Co-owner",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Product Sales",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Actuarial Analysis",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Appraiser",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales Operations",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Audit",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Proprietor",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Bookkeeping",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Claims Review/Adjusting",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Treasurer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Financial Analysis/Research",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Actuarial Analysis",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Appraiser",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Audit",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Bookkeeping",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Risk Management/Compliance",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Claims Review/Adjusting",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Financial Analysis/Research",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Risk Management/Compliance",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Agent/Broker",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Business Development",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Field Service",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - General Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - International Sales",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Media/Advertising Sales",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Product Sales",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales Operations",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Telesales",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Controller",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Analysis/Research",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Control",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Cyber Security/Information Security",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Software Development",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "standard"
        }
      ],
      "precision": [
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Automotive",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Agriculture",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Automotive",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Automotive",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Owner/Co-owner",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Property & Casualty",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Property & Casualty",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Managing Partner",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Director",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Property & Casualty",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Proprietor",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Owner",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Owner",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Agent/Broker",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Business Development",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Field Service",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Owner/Co-owner",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Actuarial Analysis",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Product Sales",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Appraiser",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Audit",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales Operations",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Proprietor",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Bookkeeping",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Claims Review/Adjusting",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Treasurer",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Financial Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Actuarial Analysis",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Appraiser",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Audit",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Bookkeeping",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Risk Management/Compliance",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Claims Review/Adjusting",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Financial Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Risk Management/Compliance",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Agent/Broker",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Business Development",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Field Service",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - International Sales",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Media/Advertising Sales",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Product Sales",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales Operations",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Telesales",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Controller",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Control",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Planning/Advising",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Cyber Security/Information Security",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Software Development",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "optional"
        }
      ],
      "targeted": [
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Automotive",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Agriculture",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Automotive",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Automotive",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Property & Casualty",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Owner/Co-owner",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Managing Partner",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Property & Casualty",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Property & Casualty",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Proprietor",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Level",
          "label": "Contractor",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Contractor",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Agent/Broker",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Business Development",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Field Service",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Owner/Co-owner",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Product Sales",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Actuarial Analysis",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Appraiser",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales Operations",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Audit",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Bookkeeping",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Proprietor",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Claims Review/Adjusting",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Financial Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Executives - Treasurer",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Actuarial Analysis",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Appraiser",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Audit",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Risk Management/Compliance",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Bookkeeping",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Claims Review/Adjusting",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Insurance - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Financial Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Policy Underwriting",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Risk Management/Compliance",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Insurance - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Agent/Broker",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Business Development",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Field Service",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - International Sales",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Media/Advertising Sales",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Product Sales",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales Operations",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Sales - Telesales",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Controller",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Control",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Planning/Advising",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Cyber Security/Information Security",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Software Development",
          "type": "optional"
        },
        {
          "subAudience": "Entire PC360 Audience",
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "optional"
        }
      ]
    }
  },
  {
    "id": "CUT",
    "name": "Credit Union Times",
    "isCustom": null,
    "logoUrl": "https://www.cutimes.com/_nuxt/img/logo-header-credit-union-times.f39782f.png",
    "subAudiences": null,
    "subAudienceCombos": {},
    "accountReps": [
      {
        "name": "Stacy Barrett",
        "email": "stacy.barrett@arc-network.com"
      }
    ],
    "customQuestions": {
      "precision": [
        {
          "id": "cu-bank-asset-size",
          "question": "Credit Union / Bank Asset Size?",
          "type": "optional",
          "subAudience": null,
          "standard": [
            "More than $10 billion (accepted)",
            "$4 billion - <$10 billion (accepted)",
            "$2 billion - <$4 billion (accepted)",
            "$1 billion - <$2 billion (accepted)",
            "$500 million - <$1 billion (accepted)",
            "$250 million - <$500 million (not accepted)",
            "Less than $250 million (not accepted)"
          ],
          "optional": []
        }
      ],
      "targeted": [
        {
          "id": "cu-bank-asset-size",
          "question": "Credit Union / Bank Asset Size?",
          "type": "optional",
          "subAudience": null,
          "standard": [
            "More than $10 billion (accepted)",
            "$4 billion - <$10 billion (accepted)",
            "$2 billion - <$4 billion (accepted)",
            "$1 billion - <$2 billion (accepted)",
            "$500 million - <$1 billion (accepted)",
            "$250 million - <$500 million (accepted)",
            "Less than $250 million (not accepted)"
          ],
          "optional": []
        }
      ]
    },
    "filterOverrides": {
      "reach": [
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Analyst",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Accounting",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Banking",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Audit",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Banking",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Credit Unions",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Branch Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial (General)",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Broker",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Credit",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Investment Banking/Venture",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Lending",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Mergers/Acquisitions",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Teller",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Title/Escrow",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Underwriter",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Board Member",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Accounting",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Accounts Payable/Receivable",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Analyst",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Audit",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Bookkeeping",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Collections",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Controller",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Corporate",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Credit Review/Analysis",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Analysis/Research",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Control",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Planning/Advising",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Product Sales/Brokerage",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Fund Accounting",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Procurement",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Risk Management/Compliance",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Securities Analysis/Research",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Tax Assessment/Collections",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Affiliate Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Brand/Product Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Channel Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Content Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Copy Writing/Editing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Corporate Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Demand/Lead Generation",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Digital Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Direct Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Events/Promotional Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Fundraising",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Investor/Public/Media Relations",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Market Research",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Marketing Communications",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Marketing Operations",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Marketing Production/Traffic",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Media Planning/Buying",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Product Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Public Relations",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - SEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Social Media",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Telemarketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Visual/Display Merchandising",
          "type": "standard"
        }
      ],
      "precision": [
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Analyst",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Accounting",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Audit",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Banking",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Banking",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Credit Unions",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial (General)",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Branch Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Broker",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Credit",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Investment Banking/Venture",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Lending",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Mergers/Acquisitions",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Owner",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Teller",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Title/Escrow",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Underwriter",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Board Member",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Accounting",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Accounts Payable/Receivable",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Analyst",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Audit",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Bookkeeping",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Collections",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Controller",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Corporate",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Credit Review/Analysis",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Control",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Planning/Advising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Product Sales/Brokerage",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Fund Accounting",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Procurement",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Risk Management/Compliance",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Securities Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Tax Assessment/Collections",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Affiliate Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Brand/Product Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Channel Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Content Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Copy Writing/Editing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Corporate Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Demand/Lead Generation",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Digital Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Direct Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Events/Promotional Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Fundraising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Investor/Public/Media Relations",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Market Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Marketing Communications",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Marketing Operations",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Marketing Production/Traffic",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Media Planning/Buying",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Product Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Public Relations",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - SEO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Social Media",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Telemarketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Visual/Display Merchandising",
          "type": "optional"
        }
      ],
      "targeted": [
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Analyst",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Accounting",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Banking",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Audit",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Banking",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Credit Unions",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial (General)",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Branch Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Broker",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Credit",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Investment Banking/Venture",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Lending",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Mergers/Acquisitions",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Owner",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Teller",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Title/Escrow",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Underwriter",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Board Member",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Accounting",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Accounts Payable/Receivable",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Analyst",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Audit",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Bookkeeping",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Collections",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Controller",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Corporate",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Credit Review/Analysis",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Control",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Planning/Advising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Product Sales/Brokerage",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Fund Accounting",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Procurement",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Risk Management/Compliance",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Securities Analysis/Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Tax Assessment/Collections",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Affiliate Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Brand/Product Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Channel Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Content Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Copy Writing/Editing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Corporate Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Demand/Lead Generation",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Digital Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Direct Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Events/Promotional Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Fundraising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Investor/Public/Media Relations",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Market Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Marketing Communications",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Marketing Operations",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Marketing Production/Traffic",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Media Planning/Buying",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Product Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Public Relations",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - SEO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Social Media",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Telemarketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Visual/Display Merchandising",
          "type": "optional"
        }
      ]
    }
  },
  {
    "id": "GST",
    "name": "GlobeSt.",
    "isCustom": null,
    "logoUrl": "https://www.globest.com/_nuxt/img/globest-blue.743e64e.png",
    "subAudiences": null,
    "subAudienceCombos": {},
    "accountReps": [
      {
        "name": "Eric Gordon",
        "email": "eric.gordon@arc-network.com"
      },
      {
        "name": "Elizabeth Ames",
        "email": "elizabeth.ames@arc-network.com"
      }
    ],
    "customQuestions": {},
    "filterOverrides": {
      "reach": [
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Commercial Real Estate",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Credit",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Mortgage",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Lending",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Underwriter",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Property Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Real Estate (General)",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Residential Real Estate",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Contractor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - General Counsel",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Managing Partner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Owner/Co-owner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Proprietor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Product Sales/Brokerage",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Investment Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Contracts Administration",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - In-House Attorney",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Law Firm Attorney",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Legal Operations",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Real Estate",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Regulatory/Compliance",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Agent/Broker",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Appraising",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Investor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Leasing/Acquisition",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Legal",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Property Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Title Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Underwriter",
          "type": "standard"
        }
      ],
      "precision": [
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Commercial Real Estate",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Credit",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Lending",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Mortgage",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Underwriter",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Property Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Real Estate (General)",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Residential Real Estate",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - General Counsel",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Managing Partner",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Owner/Co-owner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Proprietor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Product Sales/Brokerage",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Contracts Administration",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - In-House Attorney",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Law Firm Attorney",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Legal Operations",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Real Estate",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Regulatory/Compliance",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Agent/Broker",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Appraising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Investor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Leasing/Acquisition",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Legal",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Property Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Title Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Underwriter",
          "type": "optional"
        }
      ],
      "targeted": [
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Commercial Real Estate",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Credit",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Lending",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Mortgage",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Property Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Banking/Mortgage - Underwriter",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Real Estate (General)",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate - Residential Real Estate",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Owner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Contractor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Office",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - General Counsel",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - General Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Managing Partner",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Owner/Co-owner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Proprietor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Financial Product Sales/Brokerage",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Investment Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Contracts Administration",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - In-House Attorney",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Law Firm Attorney",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Legal Operations",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Real Estate",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Regulatory/Compliance",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Agent/Broker",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Appraising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Investor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Leasing/Acquisition",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Legal",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Property Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Title Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Real Estate - Underwriter",
          "type": "optional"
        }
      ]
    }
  },
  {
    "id": "HRE",
    "name": "HR Executive",
    "isCustom": null,
    "logoUrl": "https://hrexecutive.com/wp-content/uploads/HR_Executive_logo_RGB.png",
    "subAudiences": null,
    "subAudienceCombos": {},
    "accountReps": [
      {
        "name": "Steve Menc",
        "email": "steve.menc@arc-network.com"
      },
      {
        "name": "Jill Schiffman",
        "email": "jill.schiffman@arc-network.com"
      }
    ],
    "customQuestions": {},
    "filterOverrides": {
      "reach": [
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Advertising/Marketing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Aerospace/Aviation",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Agriculture",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Automotive",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Communications",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Computers and Technology",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Development",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Construction",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Diversity Management/EEO/Compliance",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - HR Systems Administration",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Learning/Training",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Organizational Development",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Government",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Healthcare/Medical",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Contractor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Recruiting",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Legal",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Talent Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Manufacturing",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Wellness",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Media",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Non-Profit/Organizations",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Retail and Consumer Goods",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Service Industry",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Telecommunications",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Transportation and Logistics",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Travel/Hospitality/Entertainment",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Utility/Energy",
          "type": "standard"
        }
      ],
      "targeted": [
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Advertising/Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Aerospace/Aviation",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Agriculture",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Automotive",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Communications",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Development",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Computers and Technology",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Construction",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Diversity Management/EEO/Compliance",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - HR Systems Administration",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Learning/Training",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Government",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Organizational Development",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Healthcare/Medical",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Recruiting",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Legal",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Manufacturing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Talent Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Media",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Wellness",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Non-Profit/Organizations",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Retail and Consumer Goods",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Service Industry",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Telecommunications",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Transportation and Logistics",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Travel/Hospitality/Entertainment",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Utility/Energy",
          "type": "optional"
        }
      ],
      "precision": [
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Advertising/Marketing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Aerospace/Aviation",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Agriculture",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Automotive",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Communications",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Computers and Technology",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Development",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Diversity Management/EEO/Compliance",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Construction",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - HR Systems Administration",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Finance",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Learning/Training",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Organizational Development",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Government",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Healthcare/Medical",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Insurance",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Recruiting",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Legal",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Talent Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Manufacturing",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Wellness",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Media",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Non-Profit/Organizations",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Real Estate",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Retail and Consumer Goods",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Service Industry",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Telecommunications",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Transportation and Logistics",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Travel/Hospitality/Entertainment",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Utility/Energy",
          "type": "optional"
        }
      ]
    }
  },
  {
    "id": "DA",
    "name": "District Administration",
    "isCustom": null,
    "logoUrl": "https://districtadministration.com/wp-content/uploads/District-Administration-Logo-full-color-1-300x80.png",
    "subAudiences": null,
    "subAudienceCombos": {},
    "accountReps": [
      {
        "name": "Amanda Holsclaw",
        "email": "amanda.holsclaw@arc-network.com"
      },
      {
        "name": "Fern Sheinman",
        "email": "Fern.Sheinman@arc-network.com"
      },
      {
        "name": "Caliann Mitoulis",
        "email": "caliann.mitoulis@arc-network.com"
      }
    ],
    "customQuestions": {
      "targeted": [
        {
          "id": "da-job-title",
          "question": "What is your job title?",
          "type": "optional",
          "subAudience": null,
          "standard": [
            "Superintendent",
            "Assistant/Deputy/Associate/Regional Superintendent"
          ],
          "optional": [
            "Curriculum Manager/Director",
            "Chief Academic Officer (CAO)",
            "Business Officer",
            "Academic Affairs",
            "Student Services/Affairs",
            "College and Career Ed/CTE",
            "Counselor ",
            "Nursing",
            "School Board",
            "Professional Development",
            "Food Services",
            "Health Services"
          ]
        },
        {
          "id": "da-district-size",
          "question": "School Districts Size: # of students",
          "type": "optional",
          "subAudience": null,
          "standard": [],
          "optional": [
            "1 - 249",
            "250 - 499",
            "500-999",
            "1000-2499",
            "2500-4999",
            "5000 - 9999",
            "10,000+"
          ]
        }
      ],
      "reach": [
        {
          "id": "da-job-title",
          "question": "What is your job title?",
          "type": "optional",
          "subAudience": null,
          "standard": [
            "Superintendent",
            "Assistant/Deputy/Associate/Regional Superintendent"
          ],
          "optional": [
            "Curriculum Manager/Director",
            "Chief Academic Officer (CAO)",
            "Business Officer",
            "Academic Affairs",
            "Student Services/Affairs",
            "College and Career Ed/CTE",
            "Counselor ",
            "Nursing",
            "School Board",
            "Professional Development",
            "Food Services",
            "Health Services"
          ]
        }
      ],
      "precision": [
        {
          "id": "da-job-title",
          "question": "What is your job title?",
          "type": "optional",
          "subAudience": null,
          "standard": [
            "Superintendent",
            "Assistant/Deputy/Associate/Regional Superintendent"
          ],
          "optional": [
            "Curriculum Manager/Director",
            "Chief Academic Officer (CAO)",
            "Business Officer",
            "Academic Affairs",
            "Student Services/Affairs",
            "College and Career Ed/CTE",
            "Counselor ",
            "Nursing",
            "School Board",
            "Professional Development",
            "Food Services",
            "Health Services"
          ]
        },
        {
          "id": "da-district-size",
          "question": "School Districts Size: # of students",
          "type": "optional",
          "subAudience": null,
          "standard": [],
          "optional": [
            "1 - 249",
            "250 - 499",
            "500-999",
            "1000-2499",
            "2500-4999",
            "5000 - 9999",
            "10,000+"
          ]
        }
      ]
    },
    "filterOverrides": {
      "targeted": [
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - E-Learning",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Administrator",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Admissions/Advising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education (General)",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Archivist/Curator",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Early Childhood Care/Development",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - K-12",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Faculty",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Fitness/Sports Training",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Librarian",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Special Education",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Board Member",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Treasurer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Procurement",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Computer/Network Security",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Cyber Security/Information Security",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Data Center",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Hardware ",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Project Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Systems Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Network/Server Administration",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Storage Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Admin/Desktop Support",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Analysis",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Telecommunications & Wireless",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Computer/Electronics",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Equipment",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Facilities",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - HVAC",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Janitorial/Cleaning",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Plumbing/Pipefitting",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Car/Truck/Van/Bus Driving",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Logistics",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Vehicle Dispatch/Routing/Scheduling",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Medical and Health - Medical Practitioner",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Medical and Health - Mental Health",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Police-Law Enforcement",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Guard",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Intelligence/Analysis",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Senior Management",
          "type": "standard"
        }
      ],
      "reach": [
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - E-Learning",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Administrator",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education (General)",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Admissions/Advising",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Archivist/Curator",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Early Childhood Care/Development",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - K-12",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Faculty",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Research",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Fitness/Sports Training",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Librarian",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Special Education",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Board Member",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Treasurer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Procurement",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Computer/Network Security",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Cyber Security/Information Security",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Data Center",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Hardware ",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Project Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Systems Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Network/Server Administration",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Storage Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Admin/Desktop Support",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Analysis",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Telecommunications & Wireless",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Computer/Electronics",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Equipment",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Facilities",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - HVAC",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Janitorial/Cleaning",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Plumbing/Pipefitting",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Car/Truck/Van/Bus Driving",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Logistics",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Vehicle Dispatch/Routing/Scheduling",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Medical and Health - Medical Practitioner",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Medical and Health - Mental Health",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Police-Law Enforcement",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Guard",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Intelligence/Analysis",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Senior Management",
          "type": "standard"
        }
      ],
      "precision": [
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Administrator",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - E-Learning",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Admissions/Advising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education (General)",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Archivist/Curator",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - K-12",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Early Childhood Care/Development",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Faculty",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Fitness/Sports Training",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Librarian",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Special Education",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Board Member",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Treasurer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Procurement",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Computer/Network Security",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Cyber Security/Information Security",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Data Center",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Hardware ",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Project Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Systems Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Network/Server Administration",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Storage Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Admin/Desktop Support",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Analysis",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Telecommunications & Wireless",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Computer/Electronics",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Equipment",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Facilities",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - HVAC",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Janitorial/Cleaning",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Plumbing/Pipefitting",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Car/Truck/Van/Bus Driving",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Logistics",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Vehicle Dispatch/Routing/Scheduling",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Medical and Health - Medical Practitioner",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Medical and Health - Mental Health",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Police-Law Enforcement",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Guard",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Intelligence/Analysis",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Senior Management",
          "type": "standard"
        }
      ]
    }
  },
  {
    "id": "UB",
    "name": "University Business",
    "isCustom": null,
    "logoUrl": "https://universitybusiness.com/wp-content/uploads/2022/08/ublogo2.png",
    "subAudiences": null,
    "subAudienceCombos": {},
    "accountReps": [
      {
        "name": "Caliann Mitoulis",
        "email": "caliann.mitoulis@arc-network.com"
      },
      {
        "name": "Amanda Holsclaw",
        "email": "amanda.holsclaw@arc-network.com"
      }
    ],
    "customQuestions": {
      "precision": [
        {
          "id": "ub-job-title",
          "question": "What is your job title?",
          "type": "optional",
          "subAudience": null,
          "standard": [
            "Provost",
            "Chief Academic Officer (CAO)"
          ],
          "optional": [
            "Academic Affairs",
            "Student Services/Affairs",
            "Career Ed/CTE",
            "Professional Development",
            "Counselor ",
            "Food Services ",
            "Health Services ",
            "Registrar",
            "Dean",
            "Bursar/Financial Aid",
            "Enrollment"
          ]
        },
        {
          "id": "ub-discipline",
          "question": "Discipline? (Select all that apply)",
          "type": "optional",
          "subAudience": null,
          "standard": [],
          "optional": [
            "Agriculture/Animal Science",
            "Allied Health",
            "Business",
            "Communications",
            "Computer Science",
            "Dental Science",
            "Engineering",
            "English Language & Literature",
            "Fine & Applied Arts",
            "Foreign Languages and Literature",
            "General Studies",
            "Interdisciplinary Studies",
            "Law",
            "Mathematics",
            "Medicine",
            "Military Science",
            "Nature & Physical Science",
            "Nursing",
            "Physical Education",
            "Social Sciences",
            "Vocational Education"
          ]
        },
        {
          "id": "ub-institution-type",
          "question": "Institution Type? (select all that apply)",
          "type": "optional",
          "subAudience": null,
          "standard": [],
          "optional": [
            "Public 2 year",
            "Public 4 year",
            "Private 2 year",
            "Private 4 year",
            "Community College"
          ]
        },
        {
          "id": "ub-enrollment-size",
          "question": "Enrollment Size? (select all that apply)",
          "type": "optional",
          "subAudience": null,
          "standard": [],
          "optional": [
            "Under 500",
            "500 - 1999",
            "2000 - 4999",
            "5000+"
          ]
        }
      ],
      "targeted": [
        {
          "id": "ub-job-title",
          "question": "What is your job title?",
          "type": "optional",
          "subAudience": null,
          "standard": [
            "Provost",
            "Chief Academic Officer (CAO)"
          ],
          "optional": [
            "Academic Affairs",
            "Student Services/Affairs",
            "Career Ed/CTE",
            "Professional Development",
            "Counselor ",
            "Food Services ",
            "Health Services ",
            "Registrar",
            "Dean",
            "Bursar/Financial Aid",
            "Enrollment"
          ]
        },
        {
          "id": "ub-discipline",
          "question": "Discipline? (Select all that apply)",
          "type": "optional",
          "subAudience": null,
          "standard": [],
          "optional": [
            "Agriculture/Animal Science",
            "Allied Health",
            "Business",
            "Communications",
            "Computer Science",
            "Dental Science",
            "Engineering",
            "English Language & Literature",
            "Fine & Applied Arts",
            "Foreign Languages and Literature",
            "General Studies",
            "Interdisciplinary Studies",
            "Law",
            "Mathematics",
            "Medicine",
            "Military Science",
            "Nature & Physical Science",
            "Nursing",
            "Physical Education",
            "Social Sciences",
            "Vocational Education"
          ]
        },
        {
          "id": "ub-institution-type",
          "question": "Institution Type? (select all that apply)",
          "type": "optional",
          "subAudience": null,
          "standard": [],
          "optional": [
            "Public 2 year",
            "Public 4 year",
            "Private 2 year",
            "Private 4 year",
            "Community College"
          ]
        },
        {
          "id": "ub-enrollment-size",
          "question": "Enrollment Size? (select all that apply)",
          "type": "optional",
          "subAudience": null,
          "standard": [],
          "optional": [
            "Under 500",
            "500 - 1999",
            "2000 - 4999",
            "5000+"
          ]
        }
      ],
      "reach": [
        {
          "id": "ub-job-title",
          "question": "What is your job title?",
          "type": "optional",
          "subAudience": null,
          "standard": [
            "Provost",
            "Chief Academic Officer (CAO)"
          ],
          "optional": [
            "Academic Affairs",
            "Student Services/Affairs",
            "Career Ed/CTE",
            "Professional Development",
            "Counselor ",
            "Food Services ",
            "Health Services ",
            "Registrar",
            "Dean",
            "Bursar/Financial Aid",
            "Enrollment"
          ]
        }
      ]
    },
    "filterOverrides": {
      "targeted": [
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - E-Learning",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Business - Fundraising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Administrator",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education (General)",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Admissions/Advising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Higher Education",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Archivist/Curator",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Fitness/Sports Training",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Board Member",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chancellor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Treasurer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Procurement",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Food Services/Hospitality - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Food Services/Hospitality - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Computer/Network Security",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Cyber Security/Information Security",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Data Center",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Hardware ",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Project Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Systems Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Network/Server Administration",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Storage Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Admin/Desktop Support",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Analysis",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Telecommunications & Wireless",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Computer/Electronics",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Equipment",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Facilities",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Janitorial/Cleaning",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Car/Truck/Van/Bus Driving",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Logistics",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Medical and Health - Mental Health",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Police-Law Enforcement",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Guard",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Intelligence/Analysis",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Senior Management",
          "type": "standard"
        }
      ],
      "precision": [
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - E-Learning",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Business - Fundraising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Administrator",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education (General)",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Admissions/Advising",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Archivist/Curator",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Higher Education",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Fitness/Sports Training",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Research",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Board Member",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chancellor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Treasurer",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Procurement",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Food Services/Hospitality - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Food Services/Hospitality - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Computer/Network Security",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Cyber Security/Information Security",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Data Center",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Hardware ",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Project Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Systems Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Network/Server Administration",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Storage Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Admin/Desktop Support",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Analysis",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Telecommunications & Wireless",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Computer/Electronics",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Equipment",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Facilities",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Janitorial/Cleaning",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Car/Truck/Van/Bus Driving",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Logistics",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - General Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Medical and Health - Mental Health",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Police-Law Enforcement",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Guard",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Intelligence/Analysis",
          "type": "optional"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Senior Management",
          "type": "standard"
        }
      ],
      "reach": [
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - E-Learning",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Business - Fundraising",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Administrator",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Executive VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education (General)",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Education Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Admissions/Advising",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Archivist/Curator",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Higher Education",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Fitness/Sports Training",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Industry/Sub-Industry",
          "label": "Education - Research",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Education - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Board Member",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CCO/Chief Compliance Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Senior Employee",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CDO/Chief Data Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CFO/Chief Financial Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chancellor",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Chairman",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO/Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CIO/Chief Information Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CISO/Chief Information Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CLO/Chief Learning Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - COO/Chief Operations Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CMO/Chief Marketing Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CPO/Chief People Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CSO/Chief Security Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - CTO/Chief Technology Officer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Other CXO",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Executives - Treasurer",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Finance/Accounting - Procurement",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Food Services/Hospitality - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Food Services/Hospitality - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Business Intelligence/Data Science",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Computer/Network Security",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Cyber Security/Information Security",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Data Center",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Hardware ",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Infrastructure",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Project Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - IT Systems Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Network/Server Administration",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Storage Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Admin/Desktop Support",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - System Analysis",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "IT/Computers/Electronics - Telecommunications & Wireless",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Computer/Electronics",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Equipment",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Facilities",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Janitorial/Cleaning",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Installation/Maintenance/Repair - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Legal - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Car/Truck/Van/Bus Driving",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Logistics",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Logistics/Transportation - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Marketing - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Medical and Health - Mental Health",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - General Management",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Police-Law Enforcement",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Guard",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Security Intelligence/Analysis",
          "type": "standard"
        },
        {
          "subAudience": null,
          "category": "Job Area/Job Function",
          "label": "Security Services - Senior Management",
          "type": "standard"
        }
      ]
    }
  },
  {
    "id": "custom",
    "name": "Custom Quote",
    "isCustom": true,
    "logoUrl": null,
    "subAudiences": null,
    "subAudienceCombos": {},
    "accountReps": [],
    "customQuestions": {},
    "filterOverrides": {}
  }
];

// Canonical display order for filter categories. Categories not in this list
// (e.g. a brand-only override category, or something added later) are
// appended after these, in whatever order they first appear.
const CATEGORY_ORDER = ["Job Area/Job Function", "Job Level", "Industry/Sub-Industry", "Number of Employees"];

// Paste your automation platform's webhook URL here (Power Automate, Zapier,
// Make, etc. — any of them work identically, since this is just a POST).
// While this is empty, submissions still generate and download the Excel
// file locally exactly as before — they just won't also be forwarded.
const SUBMISSION_WEBHOOK_URL = "https://hooks.zapier.com/hooks/catch/28749932/4hqssu8/";

const MAX_CUSTOM_QUESTIONS = 2;
const CUSTOM_QUESTION_COST = 5;

// ---------------------------------------------------------------------------
// COLOR TOKENS — ARC brand: navy #0E214C (primary), teal #05B4C2 (accent,
// used specifically for checkbox selection states).
// ---------------------------------------------------------------------------
const ink = "#181A1F";
const inkSoft = "#5B6270";
const paper = "#F1F2F4";
const paperRaised = "#FFFFFF";
const line = "#DBDEE5";
const navy = "#0E214C";
const navySoft = "#E7EAF2";
const teal = "#05B4C2";
const tealSoft = "#DFF6F8";

// ---------------------------------------------------------------------------
// Excel export — builds the downloadable workbook from the submitted payload.
// ---------------------------------------------------------------------------
// Computes a package's filter sections (universal + brand overrides) as if
// targetSubAudience were the selected sub-audience — a plain, direct match,
// no combo-union logic here. Used once per "display group" below, so a
// combination sub-audience (e.g. "Brokers & HR") can show its two source
// sub-audiences as fully separate, non-merged sections.
function computeGroupSections(pkg, brand, activePkg, targetSubAudience) {
  const base = pkg.sections || {};
  const overrides = (brand && brand.filterOverrides && brand.filterOverrides[activePkg]) || [];
  const merged = {};
  Object.keys(base).forEach((cat) => {
    merged[cat] = { standard: new Set(base[cat].standard), optional: new Set(base[cat].optional) };
  });
  overrides.forEach((o) => {
    if (o.subAudience !== null && o.subAudience !== targetSubAudience) return;
    if (!merged[o.category]) merged[o.category] = { standard: new Set(), optional: new Set() };
    merged[o.category][o.type === "standard" ? "standard" : "optional"].add(o.label);
  });
  const result = {};
  Object.keys(merged).forEach((cat) => {
    result[cat] = { standard: [...merged[cat].standard], optional: [...merged[cat].optional] };
  });
  return result;
}

function sanitizeForFilename(s) {
  return String(s || "").replace(/[\\/:*?"<>|]/g, "").trim();
}

function sanitizeSheetName(s) {
  const cleaned = String(s || "Sheet").replace(/[\\/?*[\]:]/g, "-").trim();
  return cleaned.length > 31 ? cleaned.slice(0, 31) : cleaned;
}

function formatDateDDMMYY(d) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yy = String(d.getFullYear()).slice(-2);
  return `${dd}${mm}${yy}`;
}

function buildSubmissionFilename(payload) {
  const brandPart = sanitizeForFilename(payload.brand) || "Brand";
  const companyPart = sanitizeForFilename(payload.contact.company) || "Company";
  const packagePart = sanitizeForFilename(payload.package) || "Package";
  return `${brandPart}_${companyPart}_${packagePart}_${formatDateDDMMYY(new Date())}.xlsx`;
}

function buildSubmissionWorkbook(payload) {
  const wb = XLSX.utils.book_new();

  // ---- Overview ----
  const overviewRows = [["Field", "Value"]];
  overviewRows.push(["Brand", payload.brand]);
  if (payload.subAudience) overviewRows.push(["Sub-Audience", payload.subAudience]);
  overviewRows.push(["Package", payload.package]);
  if (payload.accountRep) {
    overviewRows.push(["Account Rep", payload.accountRep.name]);
    overviewRows.push(["Account Rep Email", payload.accountRep.email]);
  }
  overviewRows.push(["Submitter Name", payload.contact.name]);
  overviewRows.push(["Submitter Email", payload.contact.email]);
  overviewRows.push(["Company Name", payload.contact.company]);
  const overviewSheet = XLSX.utils.aoa_to_sheet(overviewRows);
  overviewSheet["!cols"] = [{ wch: 18 }, { wch: 36 }];
  XLSX.utils.book_append_sheet(wb, overviewSheet, "Overview");

  // ---- One sheet per filter category, in the order categories appear in
  // the payload (which already follows the canonical category order) ----
  const categoriesSeen = [];
  payload.filters.forEach((f) => {
    if (!categoriesSeen.includes(f.category)) categoriesSeen.push(f.category);
  });
  const usedSheetNames = new Set(["Overview"]);
  categoriesSeen.forEach((cat) => {
    const rows = [["Filter", "Type"]];
    const rowsForCategory = payload.filters.filter((f) => f.category === cat);
    // If any row in this category carries a group tag (e.g. "Benefits
    // Brokers" vs "Human Resources" under a combination sub-audience like
    // "Brokers & HR"), group the rows under a sub-header per group instead
    // of listing them flat — keeps the two audiences visibly distinct
    // within the same sheet, never mixed together.
    const hasGroups = rowsForCategory.some((f) => f.group);
    if (hasGroups) {
      const groupsSeen = [];
      rowsForCategory.forEach((f) => {
        if (!groupsSeen.includes(f.group)) groupsSeen.push(f.group);
      });
      groupsSeen.forEach((groupName) => {
        rows.push([`— ${groupName} —`, ""]);
        rowsForCategory
          .filter((f) => f.group === groupName)
          .forEach((f) => rows.push([f.value, f.source === "standard" ? "Standard" : "Optional"]));
      });
    } else {
      rowsForCategory.forEach((f) => rows.push([f.value, f.source === "standard" ? "Standard" : "Optional"]));
    }
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet["!cols"] = [{ wch: 50 }, { wch: 12 }];
    let sheetName = sanitizeSheetName(cat);
    let suffix = 2;
    while (usedSheetNames.has(sheetName)) {
      sheetName = sanitizeSheetName(`${cat}`.slice(0, 28)) + " " + suffix;
      suffix += 1;
    }
    usedSheetNames.add(sheetName);
    XLSX.utils.book_append_sheet(wb, sheet, sheetName);
  });

  // ---- Custom questions (required ones always included, plus any selected optional ones) ----
  if (payload.customQuestions.length > 0) {
    const rows = [["Question", "Required", "Answer", "Type"]];
    payload.customQuestions.forEach((q) => {
      q.filters.forEach((f) =>
        rows.push([q.question, q.required ? "Yes" : "No", f.value, f.source === "standard" ? "Standard" : "Optional"])
      );
    });
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet["!cols"] = [{ wch: 45 }, { wch: 10 }, { wch: 40 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, sheet, "Custom Questions");
  }

  return wb;
}

export default function PackageFormPrototype() {
  // stage: 'brand' | 'subaudience' | 'packages' | 'confirmation'
  const [stage, setStage] = useState("brand");
  const [brandId, setBrandId] = useState(null);
  const [subAudience, setSubAudience] = useState(null);
  const [customBrandName, setCustomBrandName] = useState("");

  const brand = BRANDS.find((b) => b.id === brandId) || null;
  // Treat "has sub-audiences" as "has at least one defined" — a brand
  // flagged has_subaudience=Y with zero SubAudiences rows skips straight
  // to packages rather than showing an empty, dead-end screen.
  const brandNeedsSubAudience = (b) => !!(b && b.subAudiences && b.subAudiences.length > 0);

  const selectBrand = (id) => {
    setBrandId(id);
    setSubAudience(null);
    setAccountRep("");
    const b = BRANDS.find((x) => x.id === id);
    if (b.isCustom) {
      setStage("brand");
    } else if (brandNeedsSubAudience(b)) {
      setStage("subaudience");
    } else {
      setStage("packages");
    }
  };

  const selectSubAudience = (s) => {
    setSubAudience(s);
    setStage("packages");
  };

  const goBackFromSubaudience = () => {
    setSubAudience(null);
    setStage("brand");
  };

  const goBackFromPackages = () => {
    if (brandNeedsSubAudience(brand)) {
      setStage("subaudience");
    } else {
      setBrandId(null);
      setStage("brand");
    }
  };

  // --- package + filter state --------------------------------------------
  const [activePkg, setActivePkg] = useState(PACKAGES[0].id);
  const [optionalSelections, setOptionalSelections] = useState({});
  const [selectedCustomQuestions, setSelectedCustomQuestions] = useState({});
  const [customQAnswerSelections, setCustomQAnswerSelections] = useState({});

  // ---- Contact fields + captcha, shown once a package is being viewed.
  // Not tied to brand/package selection — it's the requester's own info.
  // accountRep IS tied to brand (its options come from that brand's
  // AccountReps rows), which is why selectBrand resets it above.
  const [accountRep, setAccountRep] = useState(""); // stores the selected rep's email
  const [contactName, setContactName] = useState("");
  const [contactEmail, setContactEmail] = useState("");
  const [contactCompany, setContactCompany] = useState("");
  const [captchaChecked, setCaptchaChecked] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [submittedFilename, setSubmittedFilename] = useState(null);
  const [deliveryStatus, setDeliveryStatus] = useState(null); // null | 'sending' | 'sent' | 'error'

  const emailLooksValid = (v) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v.trim());
  const isFormValid =
    accountRep.trim().length > 0 &&
    contactName.trim().length > 0 &&
    contactEmail.trim().length > 0 &&
    emailLooksValid(contactEmail) &&
    contactCompany.trim().length > 0 &&
    captchaChecked;

  const updateContactField = (setter) => (v) => {
    setter(v);
    if (submitted) {
      setSubmitted(false);
      setSubmittedFilename(null);
      setDeliveryStatus(null);
    }
  };
  const toggleCaptcha = () => {
    setCaptchaChecked((prev) => !prev);
    if (submitted) {
      setSubmitted(false);
      setSubmittedFilename(null);
      setDeliveryStatus(null);
    }
  };

  // Full reset — used by "Start new selection" on the confirmation screen.
  const resetAll = () => {
    setBrandId(null);
    setSubAudience(null);
    setCustomBrandName("");
    setAccountRep("");
    setActivePkg(PACKAGES[0].id);
    setOptionalSelections({});
    setSelectedCustomQuestions({});
    setCustomQAnswerSelections({});
    setContactName("");
    setContactEmail("");
    setContactCompany("");
    setCaptchaChecked(false);
    setSubmitted(false);
    setSubmittedFilename(null);
    setDeliveryStatus(null);
    setStage("brand");
  };

  const pkg = PACKAGES.find((p) => p.id === activePkg);

  // Merge the package's universal filters with any brand/sub-audience-scoped
  // additions that apply given the currently selected brand + sub-audience.
  // Overrides only ever ADD to a category's standard/optional lists.
  // Category order: universal categories first (in their data order), then
  // any override-only categories appended in the order they first appear.
  // For a normal sub-audience (or none), this is a single "group". For a
  // combination sub-audience (e.g. "Brokers & HR"), this becomes two (or
  // more) fully separate groups — one per source sub-audience — each with
  // its own independent filter sections, never merged together.
  const displayGroups = useMemo(() => {
    const comboSources = (brand && brand.subAudienceCombos && brand.subAudienceCombos[subAudience]) || null;
    const groupDefs =
      comboSources && comboSources.length > 0
        ? comboSources.map((name) => ({ key: name, label: name }))
        : [{ key: subAudience || "", label: null }];
    return groupDefs.map((g) => {
      const sections = computeGroupSections(pkg, brand, activePkg, g.key);
      const keys = Object.keys(sections);
      const known = CATEGORY_ORDER.filter((c) => keys.includes(c));
      const extra = keys.filter((c) => !CATEGORY_ORDER.includes(c));
      // Custom questions scoped to this group's specific sub-audience, plus
      // any unscoped (blank subAudience) questions that apply everywhere —
      // same blank-means-universal pattern as filter overrides.
      const allQuestions = (brand && brand.customQuestions ? brand.customQuestions[activePkg] : null) || [];
      const groupQuestions = allQuestions.filter((q) => !q.subAudience || q.subAudience === g.key);
      return {
        key: g.key,
        label: g.label,
        sections,
        categoryList: [...known, ...extra],
        requiredQuestions: groupQuestions.filter((q) => q.type === "required"),
        optionalQuestions: groupQuestions.filter((q) => q.type !== "required"),
      };
    });
  }, [pkg, brand, activePkg, subAudience]);

  const isMultiGroup = displayGroups.length > 1;

  const optSelKey = (groupKey, cat) => `${brand ? brand.id : ""}:${groupKey}:${activePkg}:${cat}`;
  const questionSetKey = (groupKey) => (brand ? `${brand.id}:${groupKey}:${activePkg}` : null);
  const selectedQIdsFor = (groupKey) => {
    const key = questionSetKey(groupKey);
    return (key && selectedCustomQuestions[key]) || [];
  };

  const toggleOptional = (groupKey, category, label) => {
    const key = optSelKey(groupKey, category);
    setOptionalSelections((prev) => {
      const set = new Set(prev[key] || []);
      set.has(label) ? set.delete(label) : set.add(label);
      return { ...prev, [key]: set };
    });
  };

  const toggleCustomQuestion = (groupKey, qId) => {
    const key = questionSetKey(groupKey);
    if (!key) return;
    setSelectedCustomQuestions((prev) => {
      const current = prev[key] || [];
      let next;
      if (current.includes(qId)) {
        next = current.filter((id) => id !== qId);
      } else if (current.length < MAX_CUSTOM_QUESTIONS) {
        next = [...current, qId];
      } else {
        next = current;
      }
      return { ...prev, [key]: next };
    });
  };

  const toggleCustomQAnswer = (groupKey, qId, label) => {
    const key = `${questionSetKey(groupKey)}:${qId}`;
    setCustomQAnswerSelections((prev) => {
      const set = new Set(prev[key] || []);
      set.has(label) ? set.delete(label) : set.add(label);
      return { ...prev, [key]: set };
    });
  };

  const payload = useMemo(() => {
    if (stage !== "packages" || !brand) return null;
    const filters = [];
    const customQuestions = [];
    let optionalQuestionCount = 0;

    displayGroups.forEach((group) => {
      group.categoryList.forEach((cat) => {
        const section = group.sections[cat];
        section.standard.forEach((f) =>
          filters.push({ category: cat, value: f, source: "standard", group: isMultiGroup ? group.label : null })
        );
        (optionalSelections[optSelKey(group.key, cat)] || new Set()).forEach((f) =>
          filters.push({ category: cat, value: f, source: "optional", group: isMultiGroup ? group.label : null })
        );
      });

      const buildQuestionEntry = (qDef, isRequired) => {
        const answerKey = `${questionSetKey(group.key)}:${qDef.id}`;
        const selectedAnswers = customQAnswerSelections[answerKey] || new Set();
        const qFilters = [];
        qDef.standard.forEach((f) => qFilters.push({ value: f, source: "standard" }));
        selectedAnswers.forEach((f) => qFilters.push({ value: f, source: "optional" }));
        return { question: qDef.question, filters: qFilters, required: isRequired, group: isMultiGroup ? group.label : null };
      };

      group.requiredQuestions.forEach((qDef) => customQuestions.push(buildQuestionEntry(qDef, true)));
      const groupSelectedQIds = selectedQIdsFor(group.key);
      groupSelectedQIds.forEach((qId) => {
        const qDef = group.optionalQuestions.find((q) => q.id === qId);
        if (qDef) {
          customQuestions.push(buildQuestionEntry(qDef, false));
          optionalQuestionCount += 1;
        }
      });
    });

    const selectedRep = (brand.accountReps || []).find((r) => r.email === accountRep) || null;

    return {
      accountRep: selectedRep,
      contact: {
        name: contactName.trim(),
        email: contactEmail.trim(),
        company: contactCompany.trim(),
        captchaVerified: captchaChecked,
      },
      brand: brand.isCustom ? customBrandName || "Custom brand (unnamed)" : brand.name,
      subAudience: subAudience || null,
      package: pkg.name,
      filters,
      customQuestions,
      customQuestionCost: optionalQuestionCount * CUSTOM_QUESTION_COST,
    };
  }, [stage, pkg, displayGroups, isMultiGroup, optionalSelections, activePkg, brand, subAudience, customBrandName, accountRep, selectedCustomQuestions, customQAnswerSelections, contactName, contactEmail, contactCompany, captchaChecked]);

  return (
    <div style={{ background: paper, color: ink, minHeight: "100%", fontFamily: "'Inter', ui-sans-serif, system-ui, sans-serif" }} className="p-6 md:p-10">
      <style>{`
        .arc-form-wrap { container-type: inline-size; }
        .arc-checkbox-grid { display: flex; flex-direction: column; gap: 0.5rem; }
        @container (min-width: 480px) {
          .arc-checkbox-grid { display: grid; grid-template-columns: 1fr 1fr; column-gap: 1.5rem; row-gap: 0.5rem; }
        }
      `}</style>
      <div className="max-w-3xl mx-auto arc-form-wrap">
        {stage !== "confirmation" && (
          <div className="mb-8">
            <h1 style={{ fontSize: "1.5rem", fontWeight: 600, letterSpacing: "-0.01em" }}>Lead Gen Package Configuration</h1>
            <div style={{ marginTop: "0.5rem" }}>
              <p style={{ color: inkSoft, fontSize: "0.9rem", margin: 0 }}>1. Select a media brand</p>
              <p style={{ color: inkSoft, fontSize: "0.9rem", margin: 0 }}>2. Select and configure audience target and filters</p>
              <p style={{ color: inkSoft, fontSize: "0.9rem", margin: 0 }}>3. Submit selections</p>
            </div>
          </div>
        )}

        {/* ---------------- Screen 4: confirmation ---------------- */}
        {stage === "confirmation" && (
          <div style={{ textAlign: "center", padding: "3.5rem 1rem" }}>
            <CheckCircle2 size={48} color={navy} style={{ marginBottom: "1.25rem" }} />
            <h1 style={{ fontSize: "2rem", fontWeight: 800, color: ink, letterSpacing: "-0.01em", marginBottom: "0.9rem" }}>
              Thank you!
            </h1>
            <p style={{ fontSize: "0.95rem", color: inkSoft, maxWidth: "26rem", margin: "0 auto", lineHeight: 1.6 }}>
              We've received your selections. A copy of your selections has also started downloading for you. Please save for reference.
            </p>
            {submittedFilename && (
              <div className="flex items-center justify-center gap-1.5" style={{ fontSize: "0.8rem", color: inkSoft, marginTop: "0.85rem" }}>
                <span>
                  Downloaded <strong style={{ color: ink }}>{submittedFilename}</strong>
                </span>
                {deliveryStatus === "sent" && <CheckCircle2 size={14} color="#1F8A5C" />}
                {deliveryStatus === "error" && <XCircle size={14} color="#C23B3B" />}
              </div>
            )}
            <button
              onClick={resetAll}
              style={{
                marginTop: "2rem",
                padding: "0.8rem 1.6rem",
                borderRadius: "10px",
                border: "none",
                background: navy,
                color: "#FFFFFF",
                fontSize: "0.9rem",
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              Start new selection
            </button>
          </div>
        )}

        {/* ---------------- Screen 1: brand ---------------- */}
        {stage === "brand" && (
          <>
            <div style={{ fontWeight: 600, fontSize: "0.95rem", marginBottom: "0.75rem" }}>Select a Media Brand</div>
            <div className="grid grid-cols-2 md:grid-cols-3 gap-2 mb-6">
              {BRANDS.map((b) => {
                const active = b.id === brandId;
                return (
                  <button
                    key={b.id}
                    onClick={() => selectBrand(b.id)}
                    style={{
                      padding: "0.85rem 1rem",
                      borderRadius: "10px",
                      border: `1px solid ${active ? navy : line}`,
                      background: active ? navySoft : paperRaised,
                      color: ink,
                      textAlign: "left",
                      cursor: "pointer",
                      fontSize: "0.88rem",
                      fontWeight: active ? 600 : 500,
                    }}
                  >
                    {b.logoUrl && (
                      <img
                        src={b.logoUrl}
                        alt=""
                        style={{ height: "24px", maxWidth: "100%", objectFit: "contain", marginBottom: "0.5rem", display: "block" }}
                      />
                    )}
                    {b.name}
                  </button>
                );
              })}
            </div>

            {brand && brand.isCustom && (
              <div style={{ marginBottom: "1.5rem" }}>
                <label style={{ fontSize: "0.82rem", color: inkSoft, display: "block", marginBottom: "0.4rem" }}>
                  Brand name
                </label>
                <input
                  value={customBrandName}
                  onChange={(e) => setCustomBrandName(e.target.value)}
                  placeholder="Enter brand name"
                  style={{
                    width: "100%",
                    padding: "0.6rem 0.8rem",
                    borderRadius: "8px",
                    border: `1px solid ${line}`,
                    fontSize: "0.88rem",
                    background: paperRaised,
                    marginBottom: "0.9rem",
                  }}
                />
                <button
                  onClick={() => setStage("packages")}
                  disabled={!customBrandName.trim()}
                  style={{
                    padding: "0.6rem 1.1rem",
                    borderRadius: "8px",
                    border: "none",
                    background: customBrandName.trim() ? navy : line,
                    color: "#FFFFFF",
                    fontSize: "0.85rem",
                    fontWeight: 600,
                    cursor: customBrandName.trim() ? "pointer" : "not-allowed",
                  }}
                >
                  Continue
                </button>
              </div>
            )}
          </>
        )}

        {/* ---------------- Screen 2: sub-audience ---------------- */}
        {stage === "subaudience" && brand && (
          <>
            <button
              onClick={goBackFromSubaudience}
              className="flex items-center gap-1.5"
              style={{ background: "none", border: "none", cursor: "pointer", color: inkSoft, fontSize: "0.8rem", marginBottom: "1.25rem", padding: 0 }}
            >
              <ArrowLeft size={14} />
              {brand.name} — change brand
            </button>

            {brand.logoUrl && (
              <img src={brand.logoUrl} alt="" style={{ height: "32px", maxWidth: "220px", objectFit: "contain", marginBottom: "0.75rem" }} />
            )}
            <div style={{ fontWeight: 600, fontSize: "1.05rem", marginBottom: "0.3rem" }}>{brand.name}</div>
            <p style={{ color: inkSoft, fontSize: "0.85rem", marginBottom: "1.1rem" }}>Select audience segment to target</p>

            <div className="flex flex-col gap-2">
              {brand.subAudiences.map((s) => (
                <button
                  key={s}
                  onClick={() => selectSubAudience(s)}
                  style={{
                    padding: "0.9rem 1.1rem",
                    borderRadius: "10px",
                    border: `1px solid ${line}`,
                    background: paperRaised,
                    color: ink,
                    textAlign: "left",
                    cursor: "pointer",
                    fontSize: "0.9rem",
                    fontWeight: 500,
                  }}
                >
                  {s}
                </button>
              ))}
            </div>
          </>
        )}

        {/* ---------------- Screen 3: packages + filters ---------------- */}
        {stage === "packages" && brand && (
          <>
            <button
              onClick={goBackFromPackages}
              className="flex items-center gap-1.5"
              style={{ background: "none", border: "none", cursor: "pointer", color: inkSoft, fontSize: "0.8rem", marginBottom: "1rem", padding: 0 }}
            >
              <ArrowLeft size={14} />
              {brand.isCustom ? customBrandName || "Custom brand" : brand.name}
              {subAudience ? ` · ${subAudience}` : ""} — {brandNeedsSubAudience(brand) ? "change target audience" : "change brand"}
            </button>

            {brand.logoUrl && (
              <img src={brand.logoUrl} alt="" style={{ height: "32px", maxWidth: "220px", objectFit: "contain", marginBottom: "1rem" }} />
            )}

            <div className="flex gap-2 mb-6" role="tablist">
              {PACKAGES.map((p) => {
                const active = p.id === activePkg;
                return (
                  <button
                    key={p.id}
                    onClick={() => setActivePkg(p.id)}
                    style={{
                      flex: 1,
                      padding: "0.85rem 1rem",
                      borderRadius: "10px",
                      border: `1px solid ${active ? navy : line}`,
                      background: active ? navy : paperRaised,
                      color: active ? "#FFFFFF" : ink,
                      textAlign: "left",
                      cursor: "pointer",
                      transition: "background 120ms ease, border-color 120ms ease",
                    }}
                  >
                    <div style={{ fontWeight: 600, fontSize: "0.95rem" }}>{p.name}</div>
                    <div style={{ fontSize: "0.78rem", marginTop: "0.15rem", color: active ? "#C9D2E6" : inkSoft }}>{p.blurb}</div>
                  </button>
                );
              })}
            </div>

            {displayGroups.map((group) => (
              <React.Fragment key={group.key || "single"}>
                {isMultiGroup && (
                  <div style={{ fontWeight: 800, fontSize: "1.05rem", color: navy, marginTop: "0.5rem", marginBottom: "0.85rem" }}>
                    {group.label}
                  </div>
                )}

                {group.categoryList.length === 0 ? (
                  <div style={{ background: "#FBF0E8", border: "1px solid #E8B98C", borderRadius: "12px", padding: "1.1rem 1.25rem", marginBottom: "1.25rem", fontSize: "0.86rem", color: "#7A4A1F" }}>
                    No filters are configured yet for {brand.isCustom ? customBrandName || "this brand" : brand.name}
                    {group.label ? ` · ${group.label}` : subAudience ? ` · ${subAudience}` : ""} in the {pkg.name} package. Add rows to the Filters sheet (universal or scoped to this brand/sub-audience) to populate this screen.
                  </div>
                ) : (
                  <>
                    {/* ---- Section 1: locked, included filters ---- */}
                    <SectionBlock title={`Included in ${pkg.name} Package`} subtitle="No action needed">
                      {group.categoryList.map((cat) => {
                        const section = group.sections[cat];
                        if (section.standard.length === 0) return null;
                        return (
                          <FilterGroup key={cat} label={cat}>
                            {section.standard.map((label) => (
                              <LockedRow key={label} label={label} />
                            ))}
                          </FilterGroup>
                        );
                      })}

                      {group.requiredQuestions.length > 0 && (
                        <div>
                          <div style={{ fontWeight: 700, fontSize: "0.95rem", color: ink, marginBottom: "0.6rem" }}>
                            Required Custom Qualifying Questions
                          </div>
                          <div className="flex flex-col gap-4">
                            {group.requiredQuestions.map((q) => {
                              const answerKey = `${questionSetKey(group.key)}:${q.id}`;
                              const selectedAnswers = customQAnswerSelections[answerKey] || new Set();
                              return (
                                <div key={q.id}>
                                  <div style={{ fontSize: "0.86rem", fontWeight: 600, color: ink, marginBottom: "0.5rem" }}>
                                    {q.question}
                                  </div>
                                  <div className="flex flex-col gap-2">
                                    {q.standard.map((label) => (
                                      <LockedRow key={label} label={label} />
                                    ))}
                                    {q.optional.map((label) => (
                                      <OptionRow
                                        key={label}
                                        label={label}
                                        checked={selectedAnswers.has(label)}
                                        onToggle={() => toggleCustomQAnswer(group.key, q.id, label)}
                                      />
                                    ))}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </SectionBlock>

                    {/* ---- Section 2: optional, editable filters ---- */}
                    <SectionBlock
                      title="Add Additional Selections"
                      subtitle="If available, select anything beyond the default selections to include in your package"
                    >
                      {group.categoryList.map((cat) => {
                        const section = group.sections[cat];
                        if (section.optional.length === 0) return null;
                        const selectedOptional = optionalSelections[optSelKey(group.key, cat)] || new Set();
                        return (
                          <FilterGroup key={cat} label={cat}>
                            {section.optional.map((label) => (
                              <OptionRow key={label} label={label} checked={selectedOptional.has(label)} onToggle={() => toggleOptional(group.key, cat, label)} />
                            ))}
                          </FilterGroup>
                        );
                      })}
                    </SectionBlock>

                    {/* ---- Section 3: optional custom qualifying questions ---- */}
                    {group.optionalQuestions.length > 0 && (
                      <SectionBlock
                        title="Optional Custom Qualifying Questions"
                        subtitle={`Each additional question adds a $${CUSTOM_QUESTION_COST}/lead cost. Select up to ${MAX_CUSTOM_QUESTIONS} questions and any additional qualifying answers.`}
                      >
                        <div style={{ fontSize: "0.78rem", color: inkSoft, marginTop: "-0.6rem", marginBottom: "0.2rem" }}>
                          {selectedQIdsFor(group.key).length} of {MAX_CUSTOM_QUESTIONS} selected
                        </div>
                        {group.optionalQuestions.map((q) => {
                          const groupSelectedQIds = selectedQIdsFor(group.key);
                          const checked = groupSelectedQIds.includes(q.id);
                          const atCap = groupSelectedQIds.length >= MAX_CUSTOM_QUESTIONS && !checked;
                          const answerKey = `${questionSetKey(group.key)}:${q.id}`;
                          const selectedAnswers = customQAnswerSelections[answerKey] || new Set();
                          return (
                            <div key={q.id}>
                              <QuestionRow label={q.question} checked={checked} disabled={atCap} onToggle={() => toggleCustomQuestion(group.key, q.id)} cost={CUSTOM_QUESTION_COST} />
                              {checked && (q.standard.length > 0 || q.optional.length > 0) && (
                                <div style={{ paddingLeft: "1.9rem", marginTop: "0.5rem" }} className="flex flex-col gap-2">
                                  {q.standard.map((label) => (
                                    <LockedRow key={label} label={label} />
                                  ))}
                                  {q.optional.map((label) => (
                                    <OptionRow
                                      key={label}
                                      label={label}
                                      checked={selectedAnswers.has(label)}
                                      onToggle={() => toggleCustomQAnswer(group.key, q.id, label)}
                                    />
                                  ))}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </SectionBlock>
                    )}
                  </>
                )}
              </React.Fragment>
            ))}

            {/* ---- Contact info (required to submit) ---- */}
            <SectionBlock title="Your information" subtitle="Required to submit your selections">
              <div className="flex flex-col gap-4">
                <SelectField
                  label="Account Rep"
                  value={accountRep}
                  onChange={updateContactField(setAccountRep)}
                  placeholder={(brand.accountReps || []).length > 0 ? "Select an account rep" : "No account reps configured for this brand"}
                  options={(brand.accountReps || []).map((r) => ({ value: r.email, label: r.name }))}
                />
                <TextField label="Full name" value={contactName} onChange={updateContactField(setContactName)} placeholder="Jane Smith" />
                <TextField label="Work email" value={contactEmail} onChange={updateContactField(setContactEmail)} placeholder="jane@company.com" type="email" error={contactEmail.trim().length > 0 && !emailLooksValid(contactEmail) ? "Enter a valid email address" : null} />
                <TextField label="Company name" value={contactCompany} onChange={updateContactField(setContactCompany)} placeholder="Acme Inc." />
              </div>
            </SectionBlock>

            {/* ---- Captcha ---- */}
            <div style={{ background: paperRaised, border: `1px solid ${line}`, borderRadius: "12px", padding: "1rem 1.25rem", marginBottom: "1.25rem" }}>
              <OptionRow label="I'm not a robot" checked={captchaChecked} onToggle={toggleCaptcha} />
              <div className="flex items-center gap-1.5" style={{ marginTop: "0.6rem", paddingLeft: "1.9rem" }}>
                <ShieldCheck size={12} color={inkSoft} style={{ opacity: 0.6 }} />
                <span style={{ fontSize: "0.7rem", color: inkSoft, opacity: 0.7 }}>reCAPTCHA verification</span>
              </div>
            </div>

            {/* ---- Submit ---- */}
            <button
              onClick={async () => {
                if (!isFormValid || !payload) return;
                const wb = buildSubmissionWorkbook(payload);
                const filename = buildSubmissionFilename(payload);
                XLSX.writeFile(wb, filename);
                setSubmittedFilename(filename);
                setSubmitted(true);
                setStage("confirmation");

                if (SUBMISSION_WEBHOOK_URL) {
                  setDeliveryStatus("sending");
                  try {
                    // Upload the real binary file to our own Pages Function first
                    // (same-origin, so no CORS concerns), and get back a public URL.
                    // Zapier then fetches the file directly from that URL — a normal
                    // file download — instead of us squeezing base64 file content
                    // through the webhook, which is what kept breaking.
                    const fileBytes = XLSX.write(wb, { bookType: "xlsx", type: "array" });
                    const blob = new Blob([fileBytes], {
                      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    });
                    const uploadRes = await fetch("/api/upload", {
                      method: "POST",
                      headers: { "X-Filename": filename },
                      body: blob,
                    });
                    if (!uploadRes.ok) throw new Error("Upload to storage failed");
                    const { url: fileUrl } = await uploadRes.json();

                    const params = new URLSearchParams({
                      filename,
                      file_url: fileUrl,
                      brand: payload.brand,
                      subAudience: payload.subAudience || "",
                      package: payload.package,
                      accountRepName: payload.accountRep ? payload.accountRep.name : "",
                      accountRepEmail: payload.accountRep ? payload.accountRep.email : "",
                      contactName: payload.contact.name,
                      contactEmail: payload.contact.email,
                      contactCompany: payload.contact.company,
                    });
                    await fetch(SUBMISSION_WEBHOOK_URL, {
                      method: "POST",
                      mode: "no-cors",
                      body: params,
                    });
                    setDeliveryStatus("sent");
                  } catch (err) {
                    setDeliveryStatus("error");
                  }
                }
              }}
              disabled={!isFormValid}
              className="flex items-center justify-center gap-2"
              style={{
                width: "100%",
                padding: "0.9rem 1rem",
                borderRadius: "10px",
                border: "none",
                background: submitted ? "#1F8A5C" : isFormValid ? navy : line,
                color: isFormValid || submitted ? "#FFFFFF" : inkSoft,
                fontSize: "0.92rem",
                fontWeight: 700,
                cursor: isFormValid && !submitted ? "pointer" : "not-allowed",
                marginBottom: "1.25rem",
              }}
            >
              {submitted && <CheckCircle2 size={16} />}
              {submitted ? "Submitted" : "Submit Selections"}
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function TextField({ label, value, onChange, placeholder, type = "text", error }) {
  return (
    <div>
      <label style={{ fontSize: "0.82rem", color: inkSoft, display: "block", marginBottom: "0.4rem" }}>
        {label} <span style={{ color: "#C23B3B" }}>*</span>
      </label>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        style={{
          width: "100%",
          padding: "0.6rem 0.8rem",
          borderRadius: "8px",
          border: `1px solid ${error ? "#C23B3B" : line}`,
          fontSize: "0.88rem",
          background: paperRaised,
          color: ink,
        }}
      />
      {error && <div style={{ fontSize: "0.74rem", color: "#C23B3B", marginTop: "0.3rem" }}>{error}</div>}
    </div>
  );
}

function SelectField({ label, value, onChange, options, placeholder, helperText }) {
  return (
    <div>
      <label style={{ fontSize: "0.82rem", color: inkSoft, display: "block", marginBottom: "0.4rem" }}>
        {label} <span style={{ color: "#C23B3B" }}>*</span>
      </label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        style={{
          width: "100%",
          padding: "0.6rem 0.8rem",
          borderRadius: "8px",
          border: `1px solid ${line}`,
          fontSize: "0.88rem",
          background: paperRaised,
          color: value ? ink : inkSoft,
        }}
      >
        <option value="">{placeholder || "Select..."}</option>
        {options.map((opt) => (
          <option key={opt.value} value={opt.value}>
            {opt.label}
          </option>
        ))}
      </select>
      {helperText && <div style={{ fontSize: "0.74rem", color: inkSoft, marginTop: "0.3rem" }}>{helperText}</div>}
    </div>
  );
}

function SectionBlock({ title, subtitle, children }) {
  return (
    <div style={{ background: paperRaised, border: `1px solid ${line}`, borderRadius: "12px", padding: "1.25rem 1.4rem", marginBottom: "1.25rem" }}>
      <div style={{ fontWeight: 600, fontSize: "1rem", marginBottom: "0.2rem" }}>{title}</div>
      <div style={{ color: inkSoft, fontSize: "0.82rem", marginBottom: "1.1rem" }}>{subtitle}</div>
      <div className="flex flex-col gap-4">{children}</div>
    </div>
  );
}

function FilterGroup({ label, children }) {
  return (
    <div>
      <div style={{ fontWeight: 700, fontSize: "0.95rem", color: ink, marginBottom: "0.6rem" }}>{label}</div>
      <div className="arc-checkbox-grid">{children}</div>
    </div>
  );
}

function LockedRow({ label }) {
  return (
    <div className="flex items-center gap-2.5" style={{ padding: "0.15rem 0" }}>
      <div
        style={{
          width: "1.15rem",
          height: "1.15rem",
          borderRadius: "5px",
          background: teal,
          border: `1px solid ${teal}`,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          flexShrink: 0,
        }}
      >
        <Check size={13} color="#FFFFFF" strokeWidth={2.5} />
      </div>
      <span style={{ fontSize: "0.86rem", color: ink }}>{label}</span>
      <Lock size={12} color={inkSoft} style={{ marginLeft: "0.15rem", opacity: 0.6 }} />
    </div>
  );
}

function OptionRow({ label, checked, onToggle }) {
  return (
    <button
      onClick={onToggle}
      className="flex items-center gap-2.5"
      style={{ padding: "0.15rem 0", background: "none", border: "none", cursor: "pointer", textAlign: "left" }}
    >
      <div
        style={{
          width: "1.15rem",
          height: "1.15rem",
          borderRadius: "5px",
          background: checked ? teal : paperRaised,
          border: `1px solid ${checked ? teal : line}`,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          flexShrink: 0,
          transition: "background 100ms ease, border-color 100ms ease",
        }}
      >
        {checked && <Check size={13} color="#FFFFFF" strokeWidth={2.5} />}
      </div>
      <span style={{ fontSize: "0.86rem", color: ink }}>{label}</span>
    </button>
  );
}

function QuestionRow({ label, checked, disabled, onToggle, cost }) {
  return (
    <button
      onClick={disabled ? undefined : onToggle}
      className="flex items-center gap-2.5"
      style={{
        padding: "0.3rem 0",
        background: "none",
        border: "none",
        cursor: disabled ? "not-allowed" : "pointer",
        textAlign: "left",
        width: "100%",
        opacity: disabled ? 0.45 : 1,
      }}
    >
      <div
        style={{
          width: "1.15rem",
          height: "1.15rem",
          borderRadius: "5px",
          background: checked ? teal : paperRaised,
          border: `1px solid ${checked ? teal : line}`,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          flexShrink: 0,
          transition: "background 100ms ease, border-color 100ms ease",
        }}
      >
        {checked && <Check size={13} color="#FFFFFF" strokeWidth={2.5} />}
      </div>
      <span style={{ fontSize: "0.86rem", color: ink, flex: 1 }}>{label}</span>
      <span style={{ fontSize: "0.74rem", color: inkSoft, fontWeight: 600, flexShrink: 0 }}>+${cost}/lead</span>
    </button>
  );
}
