from dataclasses import dataclass,field
from typing import Any
@dataclass
class ActionResult:
    ok:bool
    message:str
    data:dict[str,Any]=field(default_factory=dict)
    status:str="done"
    def to_dict(self):return {"ok":self.ok,"message":self.message,"data":self.data,"status":self.status}
def waiting(message:str,checkpoint:str,data=None):
    d=dict(data or {});d["checkpoint"]=checkpoint;return ActionResult(True,message,d,"waiting_user")
