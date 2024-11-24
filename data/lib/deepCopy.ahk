class DeepCopy{
	__new(obj){
		this.fir:=obj
	}
	Call(){
		obj:=this.fir
		,this2:={base:this}
		,this2.stack:=stack:=[]
		,this2.Copied:=Copied:=Map()
		,(this3:={tgEle:0,ele:obj,base:this2}).ActAd()
		,tg:=this3.tgEle
		while stack.Length{
			act:=stack.Pop()
			act()
		}
		return tg
	}
	ActAd(){
		/* 
		ele, tgEle_
		*/
		ele:=this.ele
		act:=(
			(ele is Array and Ad(this.Act2))
			or(ele is Map and Ad(this.Act3))
			or(ele is VarRef and true)
			or(ele is Primitive
				|| Type(ele)=='Prototype')
			or Ad(this.Act1)
		)
		Ad(act){
			copied:=this.Copied
			,this.tgele:=copied.Has(ele)?(
				copied[ele]
			):(
				tgele2:=ele.Clone()
				,Copied.Set(ele,tgele2,tgele2,tgele2)
				,this.stack.Push(act.Bind(this))
				,tgele2
			)
			return act
		}
	}
	Act1(){
		/* tgEle */
		tgEle:=this.tgEle
		for k in tgEle.OwnProps(){
			ObjHasOwnProp(dsc:=tgEle.GetOwnPropDesc(k),'Value') and (
				(this2:={tgEle:0,ele:dsc.Value,base:this}).ActAd()
				,tgvl:=this2.tgEle
				,tgvl and tgEle.DefineProp(k,{value:tgvl})
			)
		}
		bas:=ObjGetBase(tgEle)
		,(this3:={ele:bas,tgEle:0,base:this}).ActAd()
		,(this3tgEle:=this3.tgEle) and ObjSetBase(tgEle,this3tgEle)
	}
	Act2(){
		/* tgEle */
		tgEle:=this.tgEle
		for i,x in tgEle{
			(this4:={tgEle:0,ele:x,base:this}).ActAd()
			,x2:=this4.tgEle
			,x2 and tgEle[i]:=x2
		}
		for k in tgEle.OwnProps(){
			ObjHasOwnProp(dsc:=tgEle.GetOwnPropDesc(k),'Value') and (
				(this2:={tgEle:0,ele:dsc.Value,base:this}).ActAd()
				,tgvl:=this2.tgEle
				,tgvl and tgEle.DefineProp(k,{value:tgvl})
			)
		}
	}
	Act3(){
		/* tgEle */
		tgEle:=this.tgEle
		for ky,x in tgEle{
			(this4:={tgEle:0,ele:ky,base:this}).ActAd()
			,ky2:=this4.tgEle
			,ky2:=(tgEle.Delete(ky),ky2) or ky
			(this5:={tgEle:0,ele:x,base:this}).ActAd()
			,x2:=this5.tgEle
			,x2:=x2 or x
			tgEle[ky2]:=x2
		}
		for k in tgEle.OwnProps(){
			ObjHasOwnProp(dsc:=tgEle.GetOwnPropDesc(k),'Value') and (
				(this2:={tgEle:0,ele:dsc.Value,base:this}).ActAd()
				,tgvl:=this2.tgEle
				,tgvl and tgEle.DefineProp(k,{value:tgvl})
			)
		}
	}
}

;Test
; cc:={c:{cc:7},mp:Map({},{}),base:{a:5}}
; cc.c.hst:=cc
; dp:=DeepCopy(cc)
; cc2:=dp()
; OutputDebug(cc==cc2)
; OutputDebug(cc.c==cc2.c)
; OutputDebug(cc.Base==cc2.Base)
; OutputDebug(cc2.c.cc)
; OutputDebug(cc2.base.a)
; cc.mp.__Enum(2)(&ky1,&x1),cc2.mp.__Enum(2)(&ky2,&x2)
; OutputDebug(ky1==ky2)
; OutputDebug(x1==x2)
; OutputDebug('test Circle reference')
; OutputDebug(cc2.c.hst==cc2)
; OutputDebug(cc2.c.hst==cc.c.hst)
; OutputDebug('another copy')
; cc3:=dp()
; OutputDebug(cc3==cc2)
