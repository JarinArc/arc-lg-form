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
    "subAudiences": [
      "Benefits Brokers",
      "Human Resources",
      "Brokers & HR"
    ],
    "customQuestions": {
      "targeted": [
        {
          "id": "group-size",
          "question": "Do you or your organization work with groups of the following size? (please select all that apply) ",
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
          "question": "Do you or your organization work with groups of the following size? (please select all that apply) ",
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
          "standard": [
            "Yes (accepted)",
            "No (not accepted)"
          ],
          "optional": []
        }
      ]
    },
    "filterOverrides": {
      "reach": [
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Customer Support - Client Services/Account Management",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO - Chief Human Resource Officer",
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
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
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
          "label": "Executives - CLO - Chief Learning Officer",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Aerospace/Aviation",
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
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
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
          "category": "Job Area/Job Function",
          "label": "Executives - CPO - Chief People Officer",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
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
          "category": "Job Area/Job Function",
          "label": "Executives - President",
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
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
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
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
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
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
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
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
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
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
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
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Director",
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
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Development",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Employee Benefits Broker/Consultant",
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
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Construction",
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
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
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
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
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
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
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
          "category": "Number of Employees",
          "label": "250 - 499",
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
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - HR Systems Administration",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "100 - 249",
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
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "50 - 99",
          "type": "standard"
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
          "category": "Job Area/Job Function",
          "label": "Human Resources - Learning/Training",
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
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
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
          "category": "Job Level",
          "label": "Consultant",
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
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
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
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Contractor",
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
          "category": "Job Level",
          "label": "Contractor",
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
          "category": "Job Area/Job Function",
          "label": "Sales - Agent - Broker",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Insurance",
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
          "category": "Number of Employees",
          "label": "5 - 9",
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
          "category": "Number of Employees",
          "label": "5 - 9",
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
          "category": "Industry/Sub-Industry",
          "label": "Legal",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
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
      "precision": [
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Customer Support - Client Services/Account Management",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
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
          "label": "Finance - Financial Services",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "50,000+",
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
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO - Chief Human Resource Officer",
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
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
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
          "label": "Executives - CLO - Chief Learning Officer",
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
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Aerospace/Aviation",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO - Chief People Officer",
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
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Agriculture",
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
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
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
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Automotive",
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
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
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
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
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
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
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
          "category": "Industry/Sub-Industry",
          "label": "Computers and Technology",
          "type": "optional"
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
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
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
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "500 - 999",
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
          "category": "Job Level",
          "label": "Senior Manager",
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
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
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
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "100 - 249",
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
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Education",
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
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "100 - 249",
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
          "category": "Industry/Sub-Industry",
          "label": "Finance",
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
          "category": "Job Level",
          "label": "Individual Contributor",
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
          "category": "Job Level",
          "label": "Owner",
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
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Organizational Development",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Consultant",
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
          "category": "Job Level",
          "label": "Consultant",
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
          "label": "Contractor",
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
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Healthcare/Medical",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Agent - Broker",
          "type": "optional"
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
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Insurance",
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
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
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
          "category": "Job Area/Job Function",
          "label": "Human Resources - Wellness",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Media",
          "type": "optional"
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
      "targeted": [
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO - Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
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
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Customer Support - Client Services/Account Management",
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
          "label": "Executives - CLO - Chief Learning Officer",
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
          "category": "Job Area/Job Function",
          "label": "Executives - CPO - Chief People Officer",
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
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Agriculture",
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
          "category": "Job Area/Job Function",
          "label": "Executives - President",
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
          "subAudience": "Benefits Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
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
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
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
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Communications",
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
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
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
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Director",
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
          "subAudience": "Benefits Brokers",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Senior Manager",
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
          "category": "Job Level",
          "label": "Senior Manager",
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
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
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
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
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
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
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
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
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
          "category": "Number of Employees",
          "label": "100 - 249",
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
          "category": "Number of Employees",
          "label": "50 - 99",
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
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Finance",
          "type": "optional"
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
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": "Human Resources",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Organizational Development",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Government",
          "type": "optional"
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
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Consultant",
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
          "category": "Industry/Sub-Industry",
          "label": "Healthcare/Medical",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Agent - Broker",
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
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": "Benefits Brokers",
          "category": "Job Level",
          "label": "Contractor",
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
          "category": "Number of Employees",
          "label": "1 - 4",
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
          "category": "Job Area/Job Function",
          "label": "Human Resources - Wellness",
          "type": "standard"
        },
        {
          "subAudience": "Human Resources",
          "category": "Industry/Sub-Industry",
          "label": "Media",
          "type": "optional"
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
    },
    "logoUrl": "https://www.benefitspro.com/_nuxt/img/bpro-blue.0a30307.png"
  },
  {
    "id": "TA",
    "name": "ThinkAdvisor",
    "isCustom": null,
    "subAudiences": null,
    "customQuestions": {},
    "filterOverrides": {},
    "logoUrl": "https://www.thinkadvisor.com/_nuxt/img/ALM_ThinkAdvisor.84d9263.png"
  },
  {
    "id": "PC360",
    "name": "PropertyCasualty360",
    "isCustom": null,
    "subAudiences": [
      "Entire PC360 Audience",
      "Agents & Brokers",
      "Carriers",
      "Claims"
    ],
    "customQuestions": {
      "professional": [
        {
          "id": "patient-contact",
          "question": "Will this role involve direct patient contact?",
          "standard": [],
          "optional": []
        },
        {
          "id": "prod-access",
          "question": "Will this role have production-system access?",
          "standard": [],
          "optional": []
        }
      ],
      "enterprise": [
        {
          "id": "patient-contact",
          "question": "Will this role involve direct patient contact?",
          "standard": [],
          "optional": []
        },
        {
          "id": "finra",
          "question": "Is this role subject to FINRA registration?",
          "standard": [],
          "optional": []
        },
        {
          "id": "controlled-substances",
          "question": "Will this role require access to controlled substances?",
          "standard": [],
          "optional": []
        }
      ],
      "essential": [
        {
          "id": "commercial-vehicle",
          "question": "Does this role require operating a commercial vehicle?",
          "standard": [],
          "optional": []
        },
        {
          "id": "jobsite-safety",
          "question": "Does this role require jobsite safety certification?",
          "standard": [],
          "optional": []
        }
      ]
    },
    "filterOverrides": {
      "precision": [
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Customer Support - Client Services/Account Management",
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
          "label": "Finance - Financial Services",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Advertising/Marketing",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO - Chief Human Resource Officer",
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
          "label": "Executives - CEO",
          "type": "optional"
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
          "label": "20,000 - 49,999",
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
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Aerospace/Aviation",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO - Chief Learning Officer",
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
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Agriculture",
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
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO - Chief People Officer",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
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
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Automotive",
          "type": "optional"
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
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "VP",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
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
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
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
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Communications",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Employee Benefits Broker/Consultant",
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
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Computers and Technology",
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
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Development",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
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
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Diversity Management/EEO/Compliance",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "500 - 999",
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
          "category": "Industry/Sub-Industry",
          "label": "Construction",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Manager",
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
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Manager",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Education",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Insurance Brokerage",
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
          "label": "Human Resources - HR Systems Administration",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
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
          "label": "50 - 99",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Owner",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Learning/Training",
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
          "category": "Industry/Sub-Industry",
          "label": "Finance",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Government",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "InsuranceSenior Management",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Organizational Development",
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
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "10 - 24",
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
          "category": "Job Area/Job Function",
          "label": "Sales - Agent - Broker",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Healthcare/Medical",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Recruiting",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Legal",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Talent Management",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Manufacturing",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Media",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Wellness",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Non-Profit/Organizations",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Real Estate",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Retail and Consumer Goods",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Service Industry",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Telecommunications",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Transportation and Logistics",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Travel/Hospitality/Entertainment",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Utility/Energy",
          "type": "optional"
        }
      ],
      "targeted": [
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
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
          "category": "Job Area/Job Function",
          "label": "Customer Support - Client Services/Account Management",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "50,000+",
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
          "label": "Advertising/Marketing",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO - Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Aerospace/Aviation",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
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
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO - Chief Learning Officer",
          "type": "standard"
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
          "label": "Agriculture",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - President",
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
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Insurance General",
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
          "category": "Job Level",
          "label": "Senior VP",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO - Chief People Officer",
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
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Principal",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "optional"
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
          "category": "Industry/Sub-Industry",
          "label": "Automotive",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "2,500 - 4,999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
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
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Communications",
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
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Development",
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
          "label": "Computers and Technology",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Director",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Employee Benefits Broker/Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Construction",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "500 - 999",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Diversity Management/EEO/Compliance",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "optional"
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
          "label": "Manager",
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
          "category": "Job Area/Job Function",
          "label": "Insurance - Independent Insurance Agent",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
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
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
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
          "label": "Human Resources - HR Systems Administration",
          "type": "standard"
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
          "category": "Industry/Sub-Industry",
          "label": "Education",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Finance",
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
          "label": "Individual Contributor",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Learning/Training",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - Planning/Advising",
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
          "label": "Owner",
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
          "category": "Job Area/Job Function",
          "label": "InsuranceSenior Management",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Consultant",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Organizational Development",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Government",
          "type": "optional"
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
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Contractor",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Healthcare/Medical",
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
          "category": "Job Area/Job Function",
          "label": "Sales - Agent - Broker",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Recruiting",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "optional"
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
          "label": "Sales - Sales",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Legal",
          "type": "optional"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Senior Management",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Manufacturing",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Talent Management",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Media",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Wellness",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Non-Profit/Organizations",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Real Estate",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Retail and Consumer Goods",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Service Industry",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Telecommunications",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Transportation and Logistics",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Travel/Hospitality/Entertainment",
          "type": "optional"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Utility/Energy",
          "type": "optional"
        }
      ],
      "reach": [
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "C-Level",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Finance - Financial Services",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Customer Support - Client Services/Account Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Advertising/Marketing",
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
          "category": "Number of Employees",
          "label": "50,000+",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CHRO - Chief Human Resource Officer",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - CEO",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CLO - Chief Learning Officer",
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
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Health",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "20,000 - 49,999",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Aerospace/Aviation",
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
          "label": "20,000 - 49,999",
          "type": "standard"
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
          "category": "Job Area/Job Function",
          "label": "Executives - President",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Executives - CPO - Chief People Officer",
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
          "category": "Number of Employees",
          "label": "10,000 - 19,999",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Agriculture",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Compensation/Benefit Policy",
          "type": "standard"
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
          "label": "Automotive",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Industry/Sub-Industry",
          "label": "Insurance - Life",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "5,000 - 9,999",
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
          "category": "Job Level",
          "label": "VP",
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
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Executives - Partner",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Director",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Communications",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Biotech and Pharmaceuticals",
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
          "category": "Industry/Sub-Industry",
          "label": "Computers and Technology",
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
          "category": "Job Level",
          "label": "Director",
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
          "label": "Insurance - Employee Benefits Broker/Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Corporate Development",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "1,000 - 2,499",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Diversity Management/EEO/Compliance",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Senior Manager",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Insurance - General Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Construction",
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
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "500 - 999",
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
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - General Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Manager",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "250 - 499",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Corporate Services",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - HR Systems Administration",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Education",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Supervisor",
          "type": "standard"
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
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "100 - 249",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Individual Contributor",
          "type": "standard"
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
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Learning/Training",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Finance",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "50 - 99",
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
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "InsuranceSenior Management",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Organizational Development",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Consultant",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Government",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "25 - 49",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Level",
          "label": "Contractor",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Payroll/Benefits",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Healthcare/Medical",
          "type": "standard"
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
          "label": "Sales - Agent - Broker",
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
          "category": "Number of Employees",
          "label": "10 - 24",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Insurance",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Job Area/Job Function",
          "label": "Sales - Sales",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "5 - 9",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Recruiting",
          "type": "standard"
        },
        {
          "subAudience": "Agents & Brokers",
          "category": "Number of Employees",
          "label": "1 - 4",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Legal",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Senior Management",
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
          "label": "Sales - Senior Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Talent Management",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Manufacturing",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Job Area/Job Function",
          "label": "Human Resources - Wellness",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Media",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Non-Profit/Organizations",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Real Estate",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Retail and Consumer Goods",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Service Industry",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Telecommunications",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Transportation and Logistics",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Travel/Hospitality/Entertainment",
          "type": "standard"
        },
        {
          "subAudience": "Claims",
          "category": "Industry/Sub-Industry",
          "label": "Utility/Energy",
          "type": "standard"
        }
      ]
    },
    "logoUrl": "https://www.propertycasualty360.com/_nuxt/img/pc360-blue.013d9ad.png"
  },
  {
    "id": "CUT",
    "name": "Credit Union Times",
    "isCustom": null,
    "subAudiences": null,
    "customQuestions": {},
    "filterOverrides": {},
    "logoUrl": "https://www.cutimes.com/_nuxt/img/logo-header-credit-union-times.f39782f.png"
  },
  {
    "id": "GST",
    "name": "GlobeSt.",
    "isCustom": null,
    "subAudiences": null,
    "customQuestions": {},
    "filterOverrides": {},
    "logoUrl": "https://www.globest.com/_nuxt/img/globest-blue.743e64e.png"
  },
  {
    "id": "HRE",
    "name": "HR Executive",
    "isCustom": null,
    "subAudiences": null,
    "customQuestions": {},
    "filterOverrides": {},
    "logoUrl": "https://hrexecutive.com/wp-content/uploads/HR_Executive_logo_RGB.png"
  },
  {
    "id": "DA",
    "name": "District Administration",
    "isCustom": null,
    "subAudiences": null,
    "customQuestions": {},
    "filterOverrides": {},
    "logoUrl": "https://districtadministration.com/wp-content/uploads/District-Administration-Logo-full-color-1-300x80.png"
  },
  {
    "id": "UB",
    "name": "University Business",
    "isCustom": null,
    "subAudiences": null,
    "customQuestions": {},
    "filterOverrides": {},
    "logoUrl": "https://universitybusiness.com/wp-content/uploads/2022/08/ublogo2.png"
  },
  {
    "id": "custom",
    "name": "Custom Quote",
    "isCustom": true,
    "subAudiences": null,
    "customQuestions": {},
    "filterOverrides": {},
    "logoUrl": null
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
    payload.filters
      .filter((f) => f.category === cat)
      .forEach((f) => rows.push([f.value, f.source === "standard" ? "Standard" : "Optional"]));
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

  // ---- Custom questions (only if at least one was selected) ----
  if (payload.customQuestions.length > 0) {
    const rows = [["Question", "Answer", "Type"]];
    payload.customQuestions.forEach((q) => {
      q.filters.forEach((f) => rows.push([q.question, f.value, f.source === "standard" ? "Standard" : "Optional"]));
    });
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet["!cols"] = [{ wch: 45 }, { wch: 40 }, { wch: 12 }];
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
  const [contactName, setContactName] = useState("");
  const [contactEmail, setContactEmail] = useState("");
  const [contactCompany, setContactCompany] = useState("");
  const [captchaChecked, setCaptchaChecked] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [submittedFilename, setSubmittedFilename] = useState(null);
  const [deliveryStatus, setDeliveryStatus] = useState(null); // null | 'sending' | 'sent' | 'error'

  const emailLooksValid = (v) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v.trim());
  const isFormValid =
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
  const mergedSections = useMemo(() => {
    const base = pkg.sections || {};
    const overrides = (brand && brand.filterOverrides && brand.filterOverrides[activePkg]) || [];
    const merged = {};
    Object.keys(base).forEach((cat) => {
      merged[cat] = { standard: [...base[cat].standard], optional: [...base[cat].optional] };
    });
    overrides.forEach((o) => {
      if (o.subAudience !== null && o.subAudience !== subAudience) return;
      if (!merged[o.category]) merged[o.category] = { standard: [], optional: [] };
      (o.type === "standard" ? merged[o.category].standard : merged[o.category].optional).push(o.label);
    });
    return merged;
  }, [pkg, brand, activePkg, subAudience]);

  const categoryList = (() => {
    const keys = Object.keys(mergedSections);
    const known = CATEGORY_ORDER.filter((c) => keys.includes(c));
    const extra = keys.filter((c) => !CATEGORY_ORDER.includes(c));
    return [...known, ...extra];
  })();

  const optSelKey = (cat) => `${brand ? brand.id : ""}:${subAudience || ""}:${activePkg}:${cat}`;

  const availableQuestions = (brand && brand.customQuestions ? brand.customQuestions[activePkg] : null) || [];
  const questionSetKey = brand ? `${brand.id}:${activePkg}` : null;
  const selectedQIds = (questionSetKey && selectedCustomQuestions[questionSetKey]) || [];

  const toggleOptional = (category, label) => {
    const key = optSelKey(category);
    setOptionalSelections((prev) => {
      const set = new Set(prev[key] || []);
      set.has(label) ? set.delete(label) : set.add(label);
      return { ...prev, [key]: set };
    });
  };

  const toggleCustomQuestion = (qId) => {
    if (!questionSetKey) return;
    setSelectedCustomQuestions((prev) => {
      const current = prev[questionSetKey] || [];
      let next;
      if (current.includes(qId)) {
        next = current.filter((id) => id !== qId);
      } else if (current.length < MAX_CUSTOM_QUESTIONS) {
        next = [...current, qId];
      } else {
        next = current;
      }
      return { ...prev, [questionSetKey]: next };
    });
  };

  const toggleCustomQAnswer = (qId, label) => {
    const key = `${questionSetKey}:${qId}`;
    setCustomQAnswerSelections((prev) => {
      const set = new Set(prev[key] || []);
      set.has(label) ? set.delete(label) : set.add(label);
      return { ...prev, [key]: set };
    });
  };

  const payload = useMemo(() => {
    if (stage !== "packages" || !brand) return null;
    const filters = [];
    categoryList.forEach((cat) => {
      const section = mergedSections[cat];
      section.standard.forEach((f) => filters.push({ category: cat, value: f, source: "standard" }));
      (optionalSelections[optSelKey(cat)] || new Set()).forEach((f) => filters.push({ category: cat, value: f, source: "optional" }));
    });

    const customQuestions = selectedQIds.map((qId) => {
      const qDef = availableQuestions.find((q) => q.id === qId);
      const answerKey = `${questionSetKey}:${qId}`;
      const selectedAnswers = customQAnswerSelections[answerKey] || new Set();
      const qFilters = [];
      qDef.standard.forEach((f) => qFilters.push({ value: f, source: "standard" }));
      selectedAnswers.forEach((f) => qFilters.push({ value: f, source: "optional" }));
      return { question: qDef.question, filters: qFilters };
    });

    return {
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
      customQuestionCost: customQuestions.length * CUSTOM_QUESTION_COST,
    };
  }, [stage, pkg, mergedSections, categoryList, optionalSelections, activePkg, brand, subAudience, customBrandName, selectedQIds, availableQuestions, questionSetKey, customQAnswerSelections, contactName, contactEmail, contactCompany, captchaChecked]);

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

            {categoryList.length === 0 ? (
              <div style={{ background: "#FBF0E8", border: "1px solid #E8B98C", borderRadius: "12px", padding: "1.1rem 1.25rem", marginBottom: "1.25rem", fontSize: "0.86rem", color: "#7A4A1F" }}>
                No filters are configured yet for {brand.isCustom ? customBrandName || "this brand" : brand.name}
                {subAudience ? ` · ${subAudience}` : ""} in the {pkg.name} package. Add rows to the Filters sheet (universal or scoped to this brand/sub-audience) to populate this screen.
              </div>
            ) : (
              <>
                {/* ---- Section 1: locked, included filters ---- */}
                <SectionBlock title={`Included in ${pkg.name} Package`} subtitle="No action needed">
                  {categoryList.map((cat) => {
                    const section = mergedSections[cat];
                    if (section.standard.length === 0) return null;
                    return (
                      <FilterGroup key={cat} label={cat}>
                        {section.standard.map((label) => (
                          <LockedRow key={label} label={label} />
                        ))}
                      </FilterGroup>
                    );
                  })}
                </SectionBlock>

                {/* ---- Section 2: optional, editable filters ---- */}
                <SectionBlock
                  title="Add Additional Selections"
                  subtitle="If available, select anything beyond the default selections to include in your package"
                >
                  {categoryList.map((cat) => {
                    const section = mergedSections[cat];
                    if (section.optional.length === 0) return null;
                    const selectedOptional = optionalSelections[optSelKey(cat)] || new Set();
                    return (
                      <FilterGroup key={cat} label={cat}>
                        {section.optional.map((label) => (
                          <OptionRow key={label} label={label} checked={selectedOptional.has(label)} onToggle={() => toggleOptional(cat, label)} />
                        ))}
                      </FilterGroup>
                    );
                  })}
                </SectionBlock>
              </>
            )}

            {/* ---- Section 3: optional custom qualifying questions ---- */}
            {availableQuestions.length > 0 && (
              <SectionBlock
                title="Optional Custom Qualifying Questions"
                subtitle={`Each additional question adds a $${CUSTOM_QUESTION_COST}/lead cost. Select up to ${MAX_CUSTOM_QUESTIONS} questions and any additional qualifying answers.`}
              >
                <div style={{ fontSize: "0.78rem", color: inkSoft, marginTop: "-0.6rem", marginBottom: "0.2rem" }}>
                  {selectedQIds.length} of {MAX_CUSTOM_QUESTIONS} selected
                </div>
                {availableQuestions.map((q) => {
                  const checked = selectedQIds.includes(q.id);
                  const atCap = selectedQIds.length >= MAX_CUSTOM_QUESTIONS && !checked;
                  const answerKey = `${questionSetKey}:${q.id}`;
                  const selectedAnswers = customQAnswerSelections[answerKey] || new Set();
                  return (
                    <div key={q.id}>
                      <QuestionRow label={q.question} checked={checked} disabled={atCap} onToggle={() => toggleCustomQuestion(q.id)} cost={CUSTOM_QUESTION_COST} />
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
                              onToggle={() => toggleCustomQAnswer(q.id, label)}
                            />
                          ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </SectionBlock>
            )}

            {/* ---- Contact info (required to submit) ---- */}
            <SectionBlock title="Your information" subtitle="Required to submit your selections">
              <div className="flex flex-col gap-4">
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
                    const fileContentBase64 = XLSX.write(wb, { bookType: "xlsx", type: "base64" });
                    // Sent as application/x-www-form-urlencoded (via URLSearchParams) rather
                    // than JSON: it's one of the few Content-Types the browser treats as a
                    // "simple" cross-origin request (no CORS preflight, so Zapier's Catch Hook
                    // endpoint actually receives it), AND Zapier natively parses this format
                    // into individual named fields — unlike a text/plain body, which Zapier's
                    // trigger silently ignores. mode: "no-cors" means we still can't read the
                    // response, so a non-throwing fetch is treated as "sent".
                    const params = new URLSearchParams({
                      filename,
                      fileContentBase64,
                      brand: payload.brand,
                      subAudience: payload.subAudience || "",
                      package: payload.package,
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
