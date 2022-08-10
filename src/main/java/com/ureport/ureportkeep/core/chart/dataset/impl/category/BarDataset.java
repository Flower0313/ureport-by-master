/*******************************************************************************
 * Copyright 2017 Bstek
 * 
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not
 * use this file except in compliance with the License.  You may obtain a copy
 * of the License at
 * 
 *   http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  See the
 * License for the specific language governing permissions and limitations under
 * the License.
 ******************************************************************************/
package com.ureport.ureportkeep.core.chart.dataset.impl.category;


import com.ureport.ureportkeep.core.build.Context;
import com.ureport.ureportkeep.core.model.Cell;

/**
 * @author Jacky.gao
 * @since 2017年6月8日
 */
public class BarDataset extends CategoryDataset{
	@Override
	public String buildDataJson(Context context, Cell cell) {
		String datasetJson=buildDatasetJson(context, cell,null);
		StringBuilder sb=new StringBuilder();
		sb.append("{");
		String labels=getLabels();
		sb.append("\"labels\":"+labels+",");
		sb.append("\"datasets\":["+datasetJson+"]");
		sb.append("}");
		return sb.toString();
	}
	
	public String toMixJson(Context context,Cell cell,int index){
		String props="\"type\":\"bar\"";
		String datasetJson=buildDatasetJson(context, cell,props);
		return datasetJson;
	}
	@Override
	public String getType() {
		return "bar";
	}
}
