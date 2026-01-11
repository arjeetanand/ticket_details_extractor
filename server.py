from fastapi import FastAPI
from _1_ticket_automation import WeddingTicketAutomation
from _2_name_automation import VerificationWorkflow
app = FastAPI()


# @app.post("/step1-match")
# def run_step1():
#     workflow = VerificationWorkflow()
#     workflow.step1_match_and_suggest()
    
#     return {"status": "step1 completed"}

@app.post("/step2-commit")
def run_step2():
    workflow = VerificationWorkflow()
    workflow.step2_autofill_and_commit()
    return {"status": "step2 completed"}

@app.post("/ingest-and-match")
def ingest_ticket():
    automation = WeddingTicketAutomation()
    automation.run()

    workflow = VerificationWorkflow()
    workflow.step1_match_and_suggest()  # auto-suggest names

    return {"status": "ingested and matched"}

# @app.post("/run-ticket-automation")
# def run_automation():
#     automation = WeddingTicketAutomation()
#     automation.run()
#     return {"status": "Ticket automation executed"}
